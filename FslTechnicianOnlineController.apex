public with sharing class FslTechnicianOnlineController {

    // DTOs
    public class AppointmentDTO {
        @AuraEnabled public Id appointmentId;
        @AuraEnabled public String workTypeName;
        @AuraEnabled public String workOrderSubject;
        @AuraEnabled public String accountName;
        @AuraEnabled public Id accountId;
        @AuraEnabled public String contactName;
        @AuraEnabled public Id contactId;
        @AuraEnabled public String contactPhone;
        @AuraEnabled public String contactEmail;
        @AuraEnabled public Datetime schedStart;
        @AuraEnabled public Datetime schedEnd;
        @AuraEnabled public Id workOrderId;
        @AuraEnabled public String workOrderStatus;
        @AuraEnabled public String workOrderNumber;

        @AuraEnabled public String reporterContactInfo;
        @AuraEnabled public String makeModel;
        @AuraEnabled public String description;

        @AuraEnabled public String street;
        @AuraEnabled public String city;
        @AuraEnabled public String state;
        @AuraEnabled public String postalCode;
        @AuraEnabled public String country;

        @AuraEnabled public Boolean isMyAssignment;
        @AuraEnabled public Boolean isCrewAssignment;
        @AuraEnabled public Boolean allPartsEnRoute;
        @AuraEnabled public Boolean somePartsEnRoute;

        @AuraEnabled public Id serviceCrewId;
        @AuraEnabled public List<CrewMemberDTO> crewMembers;
    }

    public class CrewMemberDTO {
        @AuraEnabled public Id serviceResourceId;
        @AuraEnabled public String name;
    }

    // Unscheduled Work Orders for tray
    public class UnscheduledWorkOrderDTO {
        @AuraEnabled public Id workOrderId;
        @AuraEnabled public String workOrderNumber;
        @AuraEnabled public String subject;
        @AuraEnabled public String status;
        @AuraEnabled public Id accountId;
        @AuraEnabled public String accountName;
        @AuraEnabled public String street;
        @AuraEnabled public String city;
        @AuraEnabled public String state;
        @AuraEnabled public String postalCode;
        @AuraEnabled public String country;
    }

    public class DebugInfo {
        @AuraEnabled public Id currentUserId;
        @AuraEnabled public List<Id> serviceResourceIds;
        @AuraEnabled public List<Id> assignedResourceIds;
        @AuraEnabled public List<Id> serviceAppointmentIds;
        @AuraEnabled public List<Id> serviceCrewIds;
        @AuraEnabled public List<Id> workOrderIds;
        @AuraEnabled public Integer length;
        @AuraEnabled public String note;
    }

    public class AppointmentResult {
        @AuraEnabled public List<AppointmentDTO> appointments;
        @AuraEnabled public List<UnscheduledWorkOrderDTO> unscheduledWorkOrders;
        @AuraEnabled public DebugInfo debug;
        @AuraEnabled public String userTimeZoneId;
    }

    @AuraEnabled
    public static AppointmentResult getMyAppointmentsOnline() {
        AppointmentResult wrapper  = new AppointmentResult();
        DebugInfo dbg              = new DebugInfo();
        wrapper.debug              = dbg;

        wrapper.userTimeZoneId     = UserInfo.getTimeZone().getID();

        dbg.currentUserId          = UserInfo.getUserId();
        dbg.serviceResourceIds     = new List<Id>();
        dbg.assignedResourceIds    = new List<Id>();
        dbg.serviceAppointmentIds  = new List<Id>();
        dbg.serviceCrewIds         = new List<Id>();
        dbg.workOrderIds           = new List<Id>();
        dbg.note                   = '';

        // 1) ServiceResource for this user
        List<ServiceResource> srs = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId = :UserInfo.getUserId()
        ];

        for (ServiceResource sr : srs) {
            dbg.serviceResourceIds.add(sr.Id);
        }

        if (srs.isEmpty()) {
            dbg.note   = 'No ServiceResource found for user.';
            wrapper.appointments = new List<AppointmentDTO>();
            wrapper.unscheduledWorkOrders = new List<UnscheduledWorkOrderDTO>();
            dbg.length = 0;
            return wrapper;
        }

        Set<Id> userSrIds = new Set<Id>();
        for (ServiceResource sr : srs) {
            userSrIds.add(sr.Id);
        }

        // 2) Service crews that this user is a member of
        Set<Id> crewIds = new Set<Id>();
        for (ServiceCrewMember scm : [
            SELECT ServiceCrewId
            FROM ServiceCrewMember
            WHERE ServiceResourceId IN :userSrIds
        ]) {
            if (scm.ServiceCrewId != null) {
                crewIds.add(scm.ServiceCrewId);
                dbg.serviceCrewIds.add(scm.ServiceCrewId);
            }
        }

        // 3) AssignedResource rows tied to this user / crews
        Set<Id> saIds = new Set<Id>();
        Map<Id, Boolean> isMyBySa   = new Map<Id, Boolean>();
        Map<Id, Boolean> isCrewBySa = new Map<Id, Boolean>();
        Map<Id, Id> crewIdBySa      = new Map<Id, Id>();

        for (AssignedResource ar : [
            SELECT Id, ServiceAppointmentId, ServiceResourceId, ServiceCrewId
            FROM AssignedResource
            WHERE (ServiceResourceId IN :userSrIds
                   OR ServiceCrewId IN :crewIds)
        ]) {
            dbg.assignedResourceIds.add(ar.Id);

            if (ar.ServiceAppointmentId == null) continue;

            saIds.add(ar.ServiceAppointmentId);

            if (ar.ServiceResourceId != null && userSrIds.contains(ar.ServiceResourceId)) {
                isMyBySa.put(ar.ServiceAppointmentId, true);
            }
            if (ar.ServiceCrewId != null && crewIds.contains(ar.ServiceCrewId)) {
                isCrewBySa.put(ar.ServiceAppointmentId, true);

                if (!crewIdBySa.containsKey(ar.ServiceAppointmentId)) {
                    crewIdBySa.put(ar.ServiceAppointmentId, ar.ServiceCrewId);
                }
            }
        }

        dbg.serviceAppointmentIds.addAll(saIds);

        if (saIds.isEmpty()) {
            dbg.note   = 'No AssignedResource / ServiceAppointment rows for these resources / crews.';
            wrapper.appointments = new List<AppointmentDTO>();
            wrapper.unscheduledWorkOrders = buildUnscheduledWorkOrders(
                new Set<Id>(),
                userSrIds,
                crewIds
            );
            dbg.length = 0;
            return wrapper;
        }

        // 4) ServiceAppointments
    List<ServiceAppointment> appts = [
        SELECT Id,
            SchedStartTime,
            SchedEndTime,
            WorkType.Name,
            Subject,
            Description,
            Work_Order__c,
            Work_Order__r.Status,
            Work_Order__r.AccountId,
            Work_Order__r.Account.Name,
            Work_Order__r.ContactId,
            Work_Order__r.Contact.Name,
            Work_Order__r.Contact.Phone,
            Work_Order__r.Contact.Email,
            Work_Order__r.Make_Model__c,
            Work_Order__r.Reporter_Contact_Info__c,
            Work_Order__r.WorkOrderNumber,
            AccountId,
            Account.Name,
            Street,
            City,
            State,
            PostalCode,
            Country
            FROM ServiceAppointment
            WHERE Id IN :saIds
            AND Status != 'Canceled'
            ORDER BY SchedStartTime ASC
            LIMIT 200
        ];

        // 5) Work Orders referenced by these appointments
        Set<Id> workOrderIds = new Set<Id>();
        for (ServiceAppointment sa : appts) {
            if (sa.Work_Order__c != null) {
                workOrderIds.add(sa.Work_Order__c);
            }
        }
        dbg.workOrderIds.addAll(workOrderIds);

        // Parts info
        Map<Id, Integer> totalLinesByWo   = new Map<Id, Integer>();
        Map<Id, Integer> enRouteByWo      = new Map<Id, Integer>();

        if (!workOrderIds.isEmpty()) {
            for (WorkOrderLineItem woli : [
                SELECT Id, WorkOrderId, Line_Type__c
                FROM WorkOrderLineItem
                WHERE WorkOrderId IN :workOrderIds
            ]) {
                if (woli.WorkOrderId == null) continue;

                Integer total = totalLinesByWo.containsKey(woli.WorkOrderId)
                    ? totalLinesByWo.get(woli.WorkOrderId) : 0;
                totalLinesByWo.put(woli.WorkOrderId, total + 1);

                if (woli.Line_Type__c == 'Part En-Route') {
                    Integer shipped = enRouteByWo.containsKey(woli.WorkOrderId)
                        ? enRouteByWo.get(woli.WorkOrderId) : 0;
                    enRouteByWo.put(woli.WorkOrderId, shipped + 1);
                }
            }
        }

        // Crew members for crews on these appointments
        Set<Id> crewIdsForAppointments = new Set<Id>();
        for (Id saId : saIds) {
            if (crewIdBySa.containsKey(saId)) {
                crewIdsForAppointments.add(crewIdBySa.get(saId));
            }
        }

        Map<Id, List<CrewMemberDTO>> crewMembersByCrew = new Map<Id, List<CrewMemberDTO>>();
        if (!crewIdsForAppointments.isEmpty()) {
            for (ServiceCrewMember scm : [
                SELECT Id,
                       ServiceCrewId,
                       ServiceResourceId,
                       ServiceResource.Name
                FROM ServiceCrewMember
                WHERE ServiceCrewId IN :crewIdsForAppointments
            ]) {
                if (scm.ServiceCrewId == null || scm.ServiceResourceId == null) continue;

                List<CrewMemberDTO> listForCrew = crewMembersByCrew.get(scm.ServiceCrewId);
                if (listForCrew == null) {
                    listForCrew = new List<CrewMemberDTO>();
                    crewMembersByCrew.put(scm.ServiceCrewId, listForCrew);
                }

                CrewMemberDTO cmd = new CrewMemberDTO();
                cmd.serviceResourceId = scm.ServiceResourceId;
                cmd.name = (scm.ServiceResource != null ? scm.ServiceResource.Name : null);
                listForCrew.add(cmd);
            }
        }

        // 6) Build Appointment DTOs
        List<AppointmentDTO> results = new List<AppointmentDTO>();
        Set<Id> workOrdersWithThisTech = new Set<Id>();

        for (ServiceAppointment sa : appts) {
            AppointmentDTO dto        = new AppointmentDTO();
            dto.appointmentId         = sa.Id;
            dto.workTypeName          = (sa.WorkType == null) ? null : sa.WorkType.Name;
            dto.workOrderSubject      = sa.Subject;
                dto.workOrderNumber      = (sa.Work_Order__r != null)
                               ? sa.Work_Order__r.WorkOrderNumber
                               : null;

            dto.description           = sa.Description;
            dto.workOrderId           = sa.Work_Order__c;
            dto.workOrderStatus       = (sa.Work_Order__r != null)
                                        ? sa.Work_Order__r.Status : null;

            // Account resolution
            if (sa.AccountId != null && sa.Account != null) {
                dto.accountId   = sa.AccountId;
                dto.accountName = sa.Account.Name;
            } else if (sa.Work_Order__r != null && sa.Work_Order__r.Account != null) {
                dto.accountId   = sa.Work_Order__r.AccountId;
                dto.accountName = sa.Work_Order__r.Account.Name;
            } else {
                dto.accountId   = null;
                dto.accountName = null;
            }

            dto.contactId = (sa.Work_Order__r != null) ? sa.Work_Order__r.ContactId : null;
            if (sa.Work_Order__r != null && sa.Work_Order__r.Contact != null) {
                dto.contactName   = sa.Work_Order__r.Contact.Name;
                dto.contactPhone  = sa.Work_Order__r.Contact.Phone;
                dto.contactEmail  = sa.Work_Order__r.Contact.Email;
            }

            dto.schedStart = sa.SchedStartTime;
            dto.schedEnd   = sa.SchedEndTime;

            dto.makeModel           = (sa.Work_Order__r != null)
                                      ? sa.Work_Order__r.Make_Model__c : null;
            dto.reporterContactInfo = (sa.Work_Order__r != null)
                                      ? sa.Work_Order__r.Reporter_Contact_Info__c : null;

            dto.street      = sa.Street;
            dto.city        = sa.City;
            dto.state       = sa.State;
            dto.postalCode  = sa.PostalCode;
            dto.country     = sa.Country;

            Boolean isMy   = isMyBySa.containsKey(sa.Id);
            Boolean isCrew = isCrewBySa.containsKey(sa.Id);

            dto.isCrewAssignment = isCrew;
            dto.isMyAssignment   = isMy && !isCrew;

            // Track WOs where this tech already has an appointment
            if ((isMy || isCrew) && sa.Work_Order__c != null) {
                workOrdersWithThisTech.add(sa.Work_Order__c);
            }

            dto.serviceCrewId = crewIdBySa.containsKey(sa.Id)
                ? crewIdBySa.get(sa.Id) : null;

            if (dto.serviceCrewId != null &&
                crewMembersByCrew.containsKey(dto.serviceCrewId)) {
                dto.crewMembers = crewMembersByCrew.get(dto.serviceCrewId);
            } else {
                dto.crewMembers = new List<CrewMemberDTO>();
            }

            if (sa.Work_Order__c != null &&
                totalLinesByWo.containsKey(sa.Work_Order__c)) {

                Integer total   = totalLinesByWo.get(sa.Work_Order__c);
                Integer shipped = enRouteByWo.containsKey(sa.Work_Order__c)
                    ? enRouteByWo.get(sa.Work_Order__c) : 0;

                dto.allPartsEnRoute =
                    (total != null && total > 0 && shipped != null && shipped == total);
                dto.somePartsEnRoute =
                    (total != null && shipped != null && shipped > 0 && shipped < total);
            } else {
                dto.allPartsEnRoute  = false;
                dto.somePartsEnRoute = false;
            }

            results.add(dto);
        }

        wrapper.appointments = results;
        wrapper.unscheduledWorkOrders = buildUnscheduledWorkOrders(
            workOrdersWithThisTech,
            userSrIds,
            crewIds
        );
        dbg.length = results.size();
        if (String.isBlank(dbg.note)) {
            dbg.note = 'Online fetch succeeded.';
        }

        return wrapper;
    }

    /**
     * Build the list of work orders that:
     *  - are owned by the current user
     *  - are NOT in one of the closed statuses
     *  - do NOT already have this technician on an appointment (direct or crew)
     */
    private static List<UnscheduledWorkOrderDTO> buildUnscheduledWorkOrders(
        Set<Id> workOrdersWithThisTech,
        Set<Id> userSrIds,
        Set<Id> crewIds
    ) {
        // Work Orders that are effectively "closed" from the technician perspective
        Set<String> closedStatuses = new Set<String>{
            'Completed WO',
            'Closed',
            'Cannot Complete',
            'Canceled',
            'Closed - Submitted'
        };

        List<UnscheduledWorkOrderDTO> dtos = new List<UnscheduledWorkOrderDTO>();

        // 1) Start with Work Orders owned by the current user, not closed,
        //    and not already associated to this tech or their crew
        List<WorkOrder> candidates = [
            SELECT Id,
                WorkOrderNumber,
                Subject,
                Status,
                AccountId,
                Account.Name,
                Street,
                City,
                State,
                PostalCode,
                Country
            FROM WorkOrder
            WHERE OwnerId = :UserInfo.getUserId()
            AND Status NOT IN :closedStatuses
            AND Id NOT IN :workOrdersWithThisTech
            LIMIT 200
        ];

        if (candidates.isEmpty()) {
            return dtos;
        }

        // Collect candidate Work Order Ids
        Set<Id> candidateWoIds = new Set<Id>();
        for (WorkOrder wo : candidates) {
            candidateWoIds.add(wo.Id);
        }

        // Treat these Service Appointment statuses as closed for scheduling
        Set<String> saClosedStatuses = new Set<String>{
            'Completed',
            'Completed WO',
            'Closed',
            'Cannot Complete',
            'Canceled',
            'Closed - Submitted'
        };

        // 2) Query open ServiceAppointments for those Work Orders
        Map<Id, Id> saIdToWoId = new Map<Id, Id>();

        List<ServiceAppointment> candidateSAs = [
            SELECT Id,
                Work_Order__c,
                Status
            FROM ServiceAppointment
            WHERE Work_Order__c IN :candidateWoIds
            AND Status NOT IN :saClosedStatuses
        ];

        for (ServiceAppointment sa : candidateSAs) {
            if (sa.Work_Order__c != null) {
                saIdToWoId.put(sa.Id, sa.Work_Order__c);
            }
        }

        // 3) For those ServiceAppointments, see which have any AssignedResource rows
        Set<Id> workOrdersWithAnyAssignedResource = new Set<Id>();

        if (!saIdToWoId.isEmpty()) {
            for (AssignedResource ar : [
                SELECT ServiceAppointmentId
                FROM AssignedResource
                WHERE ServiceAppointmentId IN :saIdToWoId.keySet()
            ]) {
                Id woId = saIdToWoId.get(ar.ServiceAppointmentId);
                if (woId != null) {
                    workOrdersWithAnyAssignedResource.add(woId);
                }
            }
        }

        // 4) Build DTOs only for Work Orders that:
        //    - are owned by the user
        //    - are not closed
        //    - are not already on this tech calendar (workOrdersWithThisTech)
        //    - do not have any open ServiceAppointment with any AssignedResource
        for (WorkOrder wo : candidates) {
            if (workOrdersWithAnyAssignedResource.contains(wo.Id)) {
                // Already has an open Service Appointment with assigned resources
                // so it should not appear as needs scheduling
                continue;
            }

            UnscheduledWorkOrderDTO dto = new UnscheduledWorkOrderDTO();
            dto.workOrderId     = wo.Id;
            dto.workOrderNumber = wo.WorkOrderNumber;
            dto.subject         = wo.Subject;
            dto.status          = wo.Status;
            dto.accountId       = wo.AccountId;
            dto.accountName     = (wo.Account != null ? wo.Account.Name : null);
            dto.street          = wo.Street;
            dto.city            = wo.City;
            dto.state           = wo.State;
            dto.postalCode      = wo.PostalCode;
            dto.country         = wo.Country;
            dtos.add(dto);
        }

        return dtos;
    }




    @AuraEnabled
    public static void rescheduleAppointment(Id appointmentId, Datetime newStart) {
        if (appointmentId == null || newStart == null) {
            throw new AuraHandledException('Appointment ID and new start time are required.');
        }

        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :appointmentId
            LIMIT 1
            FOR UPDATE
        ];

        Datetime oldStart = sa.SchedStartTime;
        Datetime oldEnd   = sa.SchedEndTime;

        Long durationMs;
        if (oldStart != null && oldEnd != null) {
            durationMs = oldEnd.getTime() - oldStart.getTime();
        } else {
            durationMs = 60L * 60L * 1000L; // 1 hour default
        }

        sa.SchedStartTime = newStart;
        sa.SchedEndTime   = Datetime.newInstance(newStart.getTime() + durationMs);

        update sa;
    }

    @AuraEnabled
    public static void updateAppointmentEnd(Id appointmentId, Datetime newEnd) {
        if (appointmentId == null || newEnd == null) {
            throw new AuraHandledException('Appointment ID and new end time are required.');
        }

        ServiceAppointment sa = [
            SELECT Id, SchedStartTime
            FROM ServiceAppointment
            WHERE Id = :appointmentId
            LIMIT 1
            FOR UPDATE
        ];

        Datetime startTime = sa.SchedStartTime;

        if (startTime != null && newEnd.getTime() <= startTime.getTime()) {
            // Enforce a minimum 15-minute duration
            newEnd = Datetime.newInstance(startTime.getTime() + (15 * 60 * 1000));
        }

        sa.SchedEndTime = newEnd;
        update sa;
    }

    @AuraEnabled
    public static void assignCrewAppointment(Id serviceAppointmentId, Id serviceResourceId) {
        if (serviceAppointmentId == null || serviceResourceId == null) {
            throw new AuraHandledException('Service Appointment and Service Resource are required.');
        }

        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :serviceAppointmentId
            FOR UPDATE
        ];

        Datetime originalStart = sa.SchedStartTime;
        Datetime originalEnd   = sa.SchedEndTime;

        if (originalStart == null || originalEnd == null) {
            throw new AuraHandledException(
                'This appointment is not currently scheduled. Please schedule it before assigning.'
            );
        }

        List<AssignedResource> existingArs = [
            SELECT Id
            FROM AssignedResource
            WHERE ServiceAppointmentId = :serviceAppointmentId
            FOR UPDATE
        ];

        if (!existingArs.isEmpty()) {
            delete existingArs;
        }

        ServiceAppointment saUpdate = new ServiceAppointment(
            Id = sa.Id,
            SchedStartTime = originalStart,
            SchedEndTime   = originalEnd
        );
        update saUpdate;

        AssignedResource newAr = new AssignedResource(
            ServiceAppointmentId = serviceAppointmentId,
            ServiceResourceId    = serviceResourceId
        );
        insert newAr;
    }

    /**
     * Create a new Service Appointment from an owned Work Order at the given start time,
     * and assign it to the current user's ServiceResource.
     */
    @AuraEnabled
    public static Id createAppointmentForWorkOrder(
        Id workOrderId,
        Datetime startDateTimeIso,
        Datetime endDateTimeIso
    ) {
        if (workOrderId == null || startDateTimeIso == null) {
            throw new AuraHandledException('Work Order and start date/time are required.');
        }

        // Current user's ServiceResource
        ServiceResource sr = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId = :UserInfo.getUserId()
            LIMIT 1
        ];

        WorkOrder wo = [
            SELECT Id,
                   Subject,
                   AccountId,
                   Street,
                   City,
                   State,
                   PostalCode,
                   Country
            FROM WorkOrder
            WHERE Id = :workOrderId
            LIMIT 1
        ];

        Long durationMs = 6L * 60L * 60L * 1000L; // 6 hour default
        if (endDateTimeIso != null && endDateTimeIso.getTime() > startDateTimeIso.getTime()) {
            durationMs = endDateTimeIso.getTime() - startDateTimeIso.getTime();
        }

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c    = wo.Id,
            Subject          = wo.Subject,
            Street           = wo.Street,
            City             = wo.City,
            State            = wo.State,
            PostalCode       = wo.PostalCode,
            Country          = wo.Country,
            SchedStartTime   = startDateTimeIso,
            SchedEndTime     = Datetime.newInstance(startDateTimeIso.getTime() + durationMs),
            Status           = 'Scheduled'
        );
        insert sa;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = sr.Id
        );
        insert ar;

        return sa.Id;
    }
}
