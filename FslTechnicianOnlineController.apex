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
        @AuraEnabled public Boolean hasQuoteAttachment;
        @AuraEnabled public String quoteAttachmentDownloadUrl;
        @AuraEnabled public Id quoteAttachmentDocumentId;
        @AuraEnabled public String opportunityRecordType;

        @AuraEnabled public Id serviceCrewId;
        @AuraEnabled public List<CrewMemberDTO> crewMembers;
    }

    public class CrewMemberDTO {
        @AuraEnabled public Id serviceResourceId;
        @AuraEnabled public String name;
    }

    public class ServiceResourceOption {
        @AuraEnabled public Id serviceResourceId;
        @AuraEnabled public String name;
        @AuraEnabled public Id userId;
    }

    public class ResourceAbsenceDTO {
        @AuraEnabled public Id absenceId;
        @AuraEnabled public Id resourceId;
        @AuraEnabled public Datetime start;
        @AuraEnabled public Datetime endTime;
        @AuraEnabled public String description;
        @AuraEnabled public String type;
    }

    // Unscheduled Work Orders for tray
    public class UnscheduledWorkOrderDTO {
        @AuraEnabled public Id workOrderId;
        @AuraEnabled public String workOrderNumber;
        @AuraEnabled public String subject;
        @AuraEnabled public String status;
        @AuraEnabled public Boolean hasCrewAssignment;
        @AuraEnabled public Id accountId;
        @AuraEnabled public String accountName;
        @AuraEnabled public String opportunityRecordType;
        @AuraEnabled public Boolean hasQuoteAttachment;
        @AuraEnabled public String quoteAttachmentDownloadUrl;
        @AuraEnabled public Id quoteAttachmentDocumentId;
        @AuraEnabled public String street;
        @AuraEnabled public String city;
        @AuraEnabled public String state;
        @AuraEnabled public String postalCode;
        @AuraEnabled public String country;
    }

    public class TransferRequestDTO {
        @AuraEnabled public Id transferRequestId;
        @AuraEnabled public Id workOrderId;
        @AuraEnabled public String workOrderNumber;
        @AuraEnabled public String workOrderSubject;
        @AuraEnabled public String accountName;
        @AuraEnabled public String street;
        @AuraEnabled public String city;
        @AuraEnabled public String state;
        @AuraEnabled public String postalCode;
        @AuraEnabled public String country;
        @AuraEnabled public String requestedByName;
        @AuraEnabled public String targetName;
        @AuraEnabled public Datetime acceptedOn;
        @AuraEnabled public Datetime rejectedOn;
        @AuraEnabled public String rejectionReason;
    }

    public class DebugInfo {
        @AuraEnabled public Id currentUserId;
        @AuraEnabled public List<Id> serviceResourceIds;
        @AuraEnabled public List<Id> absenceResourceIds;
        @AuraEnabled public List<Id> assignedResourceIds;
        @AuraEnabled public List<Id> serviceAppointmentIds;
        @AuraEnabled public List<Id> serviceCrewIds;
        @AuraEnabled public List<Id> workOrderIds;
        @AuraEnabled public Integer length;
        @AuraEnabled public Integer absenceCount;
        @AuraEnabled public String note;
    }

    public class AppointmentResult {
        @AuraEnabled public List<AppointmentDTO> appointments;
        @AuraEnabled public List<UnscheduledWorkOrderDTO> unscheduledWorkOrders;
        @AuraEnabled public List<TransferRequestDTO> transferRequests;
        @AuraEnabled public List<TransferRequestDTO> submittedTransferRequests;
        @AuraEnabled public List<ResourceAbsenceDTO> absences;
        @AuraEnabled public DebugInfo debug;
        @AuraEnabled public String userTimeZoneId;
        @AuraEnabled public Boolean isManager;
        @AuraEnabled public List<UserPickerOption> teamMembers;
        @AuraEnabled public Id viewingUserId;
    }

    public class UserPickerOption {
        @AuraEnabled public Id userId;
        @AuraEnabled public String name;
    }

    @AuraEnabled
    public static AppointmentResult getMyAppointmentsOnline(Id targetUserId) {
        AppointmentResult wrapper  = new AppointmentResult();
        DebugInfo dbg              = new DebugInfo();
        wrapper.debug              = dbg;

        Id effectiveUserId = (targetUserId != null)
            ? targetUserId
            : UserInfo.getUserId();
        Set<Id> srLookupIds = new Set<Id>();
        if (effectiveUserId != null) {
            srLookupIds.add(effectiveUserId);
        }
        Id viewerUserId = UserInfo.getUserId();

        wrapper.userTimeZoneId     = UserInfo.getTimeZone().getID();

        dbg.currentUserId          = UserInfo.getUserId();
        dbg.serviceResourceIds     = new List<Id>();
        dbg.absenceResourceIds     = new List<Id>();
        dbg.assignedResourceIds    = new List<Id>();
        dbg.serviceAppointmentIds  = new List<Id>();
        dbg.serviceCrewIds         = new List<Id>();
        dbg.workOrderIds           = new List<Id>();
        dbg.note                   = '';
        dbg.absenceCount           = 0;

        wrapper.teamMembers = new List<UserPickerOption>();

        for (User u : [
            SELECT Id, Name
            FROM User
            WHERE ManagerId = :UserInfo.getUserId()
            ORDER BY Name
        ]) {
            UserPickerOption opt = new UserPickerOption();
            opt.userId = u.Id;
            opt.name   = u.Name;
            wrapper.teamMembers.add(opt);
        }

        wrapper.isManager = !wrapper.teamMembers.isEmpty();
        wrapper.viewingUserId = effectiveUserId;

        // 1) ServiceResource for this user
        List<ServiceResource> srs = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId IN :srLookupIds
               OR Id IN :srLookupIds
        ];

        for (ServiceResource sr : srs) {
            dbg.serviceResourceIds.add(sr.Id);
        }

        if (srs.isEmpty()) {
            dbg.note   = 'No ServiceResource found for user.';
            wrapper.appointments = new List<AppointmentDTO>();
            wrapper.unscheduledWorkOrders = new List<UnscheduledWorkOrderDTO>();
            wrapper.absences = new List<ResourceAbsenceDTO>();
            List<Engineer_Transfer_Request__c> pendingRequests =
                getPendingTransferRequests(effectiveUserId);
            wrapper.transferRequests = buildTransferRequests(pendingRequests);
            wrapper.submittedTransferRequests =
                buildSubmittedTransferRequests(effectiveUserId);
            dbg.length = 0;
            return wrapper;
        }

        Set<Id> userSrIds = new Set<Id>();
        for (ServiceResource sr : srs) {
            userSrIds.add(sr.Id);
        }

        // ServiceResource for the *active* user; used for showing absences
        Set<Id> absenceSrIds = new Set<Id>();
        for (ServiceResource sr : [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId IN :srLookupIds
               OR Id IN :srLookupIds
        ]) {
            absenceSrIds.add(sr.Id);
        }

        wrapper.absences = new List<ResourceAbsenceDTO>();
        if (!absenceSrIds.isEmpty()) {
            // Pull a wide absence window so that future-dated time off is visible
            // even if it extends beyond the current 2-week calendar view. Managers
            // and techs commonly schedule vacation months in advance, so we include
            // two years of future absences and one year of history for context.
            Datetime windowStart = System.now().addDays(-365);
            Datetime windowEnd = System.now().addDays(365 * 2);

            for (ResourceAbsence ra : [
                SELECT Id,
                       ResourceId,
                       Start,
                       End,
                       Description,
                       Type
                FROM ResourceAbsence
                WHERE ResourceId IN :absenceSrIds
                AND End >= :windowStart
                AND Start <= :windowEnd
                ORDER BY Start ASC
            ]) {
                ResourceAbsenceDTO dto = new ResourceAbsenceDTO();
                dto.absenceId   = ra.Id;
                dto.resourceId  = ra.ResourceId;
                dto.start       = ra.Start;
                dto.endTime     = ra.End;
                dto.description = ra.Description;
                dto.type        = ra.Type;
                wrapper.absences.add(dto);
                dbg.absenceCount++;
                dbg.absenceResourceIds.add(ra.ResourceId);
            }
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

        Map<Id, String> opportunityRecordTypeByWo = new Map<Id, String>();
        Map<Id, Id> quoteDocIdByWo = new Map<Id, Id>();
        Set<Id> workOrdersWithAttachment = new Set<Id>();
        Map<Id, String> quoteDownloadByWo = new Map<Id, String>();
        Set<Id> workOrdersWithThisTech = new Set<Id>();

        if (saIds.isEmpty()) {
            dbg.note   = 'No AssignedResource / ServiceAppointment rows for these resources / crews.';
            List<Engineer_Transfer_Request__c> pendingRequests =
                getPendingTransferRequests(effectiveUserId);
            Set<Id> pendingTransferWoIds = extractWorkOrderIds(pendingRequests);
            wrapper.appointments = new List<AppointmentDTO>();
            wrapper.unscheduledWorkOrders = buildUnscheduledWorkOrders(
                workOrdersWithThisTech,
                userSrIds,
                crewIds,
                effectiveUserId,
                pendingTransferWoIds,
                workOrdersWithAttachment,
                quoteDownloadByWo,
                quoteDocIdByWo,
                opportunityRecordTypeByWo
            );
            wrapper.transferRequests = buildTransferRequests(pendingRequests);
            wrapper.submittedTransferRequests =
                buildSubmittedTransferRequests(effectiveUserId);
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
            Work_Order__r.Opportunity__c,
            Work_Order__r.Opportunity__r.RecordType.Name,
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
        Set<Id> opportunityIds = new Set<Id>();
        Map<Id, List<Id>> workOrdersByOpportunity = new Map<Id, List<Id>>();


        for (ServiceAppointment sa : appts) {
            if (sa.Work_Order__c != null) {
                workOrderIds.add(sa.Work_Order__c);

                if (sa.Work_Order__r != null && sa.Work_Order__r.Opportunity__c != null) {
                    Id oppId = sa.Work_Order__r.Opportunity__c;
                    opportunityIds.add(oppId);

                    List<Id> relatedWorkOrders = workOrdersByOpportunity.containsKey(oppId)
                        ? workOrdersByOpportunity.get(oppId)
                        : new List<Id>();
                    if (!relatedWorkOrders.contains(sa.Work_Order__c)) {
                        relatedWorkOrders.add(sa.Work_Order__c);
                    }
                    workOrdersByOpportunity.put(oppId, relatedWorkOrders);

                    if (sa.Work_Order__r.Opportunity__r != null &&
                        sa.Work_Order__r.Opportunity__r.RecordType != null) {
                        opportunityRecordTypeByWo.put(
                            sa.Work_Order__c,
                            sa.Work_Order__r.Opportunity__r.RecordType.Name
                        );
                    }
                }
            }
        }
        dbg.workOrderIds.addAll(workOrderIds);

        // Track Work Orders with an attachment (for quote tracking)


        if (!workOrderIds.isEmpty()) {
            for (ContentDocumentLink cdl : [
                SELECT LinkedEntityId,
                    ContentDocumentId,
                    ContentDocument.FileExtension,
                    ContentDocument.LatestPublishedVersionId
                FROM ContentDocumentLink
                WHERE LinkedEntityId IN :workOrderIds
            ]) {
                if (cdl.LinkedEntityId != null) {
                    workOrdersWithAttachment.add(cdl.LinkedEntityId);

                    if (cdl.ContentDocument != null &&
                        String.isNotBlank(cdl.ContentDocument.LatestPublishedVersionId) &&
                        !quoteDownloadByWo.containsKey(cdl.LinkedEntityId)) {

                        quoteDownloadByWo.put(
                            cdl.LinkedEntityId,
                            '/sfc/servlet.shepherd/document/download/' +
                                cdl.ContentDocumentId
                        );

                        if (cdl.ContentDocumentId != null) {
                            quoteDocIdByWo.put(
                                cdl.LinkedEntityId,
                                cdl.ContentDocumentId
                            );
                        }
                    }
                }
            }
        }

        if (!opportunityIds.isEmpty()) {
            for (ContentDocumentLink cdl : [
                SELECT LinkedEntityId,
                    ContentDocumentId,
                    ContentDocument.FileExtension,
                    ContentDocument.LatestPublishedVersionId
                FROM ContentDocumentLink
                WHERE LinkedEntityId IN :opportunityIds
            ]) {
                if (cdl.LinkedEntityId != null &&
                    workOrdersByOpportunity.containsKey(cdl.LinkedEntityId)) {

                    for (Id woId : workOrdersByOpportunity.get(cdl.LinkedEntityId)) {
                        workOrdersWithAttachment.add(woId);

                        if (cdl.ContentDocument != null &&
                            String.isNotBlank(cdl.ContentDocument.LatestPublishedVersionId) &&
                            !quoteDownloadByWo.containsKey(woId)) {

                            quoteDownloadByWo.put(
                                woId,
                                '/sfc/servlet.shepherd/document/download/' +
                                    cdl.ContentDocumentId
                            );

                            if (cdl.ContentDocumentId != null) {
                                quoteDocIdByWo.put(
                                    woId,
                                    cdl.ContentDocumentId
                                );
                            }
                        }
                    }
                }
            }
        }

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

            Boolean hasSchedule =
                (sa.SchedStartTime != null) || (sa.SchedEndTime != null);

            // Track WOs where this tech already has a scheduled appointment
            if ((isMy || isCrew) && hasSchedule && sa.Work_Order__c != null) {
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

            Boolean hasAttachment =
                (sa.Work_Order__c != null && workOrdersWithAttachment.contains(sa.Work_Order__c));

            dto.quoteAttachmentDownloadUrl = (sa.Work_Order__c != null &&
                quoteDownloadByWo.containsKey(sa.Work_Order__c))
                    ? quoteDownloadByWo.get(sa.Work_Order__c)
                    : null;
            dto.quoteAttachmentDocumentId = (sa.Work_Order__c != null &&
                quoteDocIdByWo.containsKey(sa.Work_Order__c))
                    ? quoteDocIdByWo.get(sa.Work_Order__c)
                    : null;

            dto.hasQuoteAttachment = hasAttachment ||
                (dto.workOrderStatus == 'Quote Attached');

            dto.opportunityRecordType = (sa.Work_Order__c != null &&
                opportunityRecordTypeByWo.containsKey(sa.Work_Order__c))
                    ? opportunityRecordTypeByWo.get(sa.Work_Order__c)
                    : null;

            results.add(dto);
        }

        List<Engineer_Transfer_Request__c> pendingRequests =
            getPendingTransferRequests(effectiveUserId);
        Set<Id> pendingTransferWoIds = extractWorkOrderIds(pendingRequests);

        wrapper.appointments = results;
        wrapper.unscheduledWorkOrders = buildUnscheduledWorkOrders(
            workOrdersWithThisTech,
            userSrIds,
            crewIds,
            effectiveUserId,
            pendingTransferWoIds,
            workOrdersWithAttachment,
            quoteDownloadByWo,
            quoteDocIdByWo,
            opportunityRecordTypeByWo
        );
        wrapper.transferRequests = buildTransferRequests(pendingRequests);
        wrapper.submittedTransferRequests =
            buildSubmittedTransferRequests(effectiveUserId);
        dbg.length = results.size();
        if (String.isBlank(dbg.note)) {
            dbg.note = 'Online fetch succeeded.';
        }

        return wrapper;
    }

    private static List<Engineer_Transfer_Request__c> getPendingTransferRequests(
        Id targetUserId
    ) {
        if (targetUserId == null) {
            return new List<Engineer_Transfer_Request__c>();
        }

        return [
            SELECT Id,
                Work_Order_to_Transfer__c,
                Work_Order_to_Transfer__r.WorkOrderNumber,
                Work_Order_to_Transfer__r.Subject,
                Work_Order_to_Transfer__r.Account.Name,
                Work_Order_to_Transfer__r.Street,
                Work_Order_to_Transfer__r.City,
                Work_Order_to_Transfer__r.State,
                Work_Order_to_Transfer__r.PostalCode,
                Work_Order_to_Transfer__r.Country,
                CreatedBy.Name,
                Transfer_Work_Order_To_Engineer__r.Name,
                Accepted_on__c,
                Date_Rejected_On__c,
                Rejected_Reason__c
            FROM Engineer_Transfer_Request__c
            WHERE Transfer_Work_Order_To_Engineer__c = :targetUserId
            AND Accepted_on__c = null
            AND Date_Rejected_On__c = null
            ORDER BY CreatedDate ASC
        ];
    }

    private static List<TransferRequestDTO> buildTransferRequests(
        List<Engineer_Transfer_Request__c> requests
    ) {
        List<TransferRequestDTO> dtos = new List<TransferRequestDTO>();

        for (Engineer_Transfer_Request__c req : requests) {
            TransferRequestDTO dto = new TransferRequestDTO();
            dto.transferRequestId = req.Id;
            dto.workOrderId = req.Work_Order_to_Transfer__c;
            dto.workOrderNumber = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.WorkOrderNumber
                : null;
            dto.workOrderSubject = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.Subject
                : null;
            dto.accountName = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.Account.Name
                : null;
            dto.street = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.Street
                : null;
            dto.city = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.City
                : null;
            dto.state = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.State
                : null;
            dto.postalCode = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.PostalCode
                : null;
            dto.country = req.Work_Order_to_Transfer__r != null
                ? req.Work_Order_to_Transfer__r.Country
                : null;
            dto.requestedByName = req.CreatedBy != null ? req.CreatedBy.Name : null;
            dto.targetName = req.Transfer_Work_Order_To_Engineer__r != null
                ? req.Transfer_Work_Order_To_Engineer__r.Name
                : null;
            dto.acceptedOn = req.Accepted_on__c;
            dto.rejectedOn = req.Date_Rejected_On__c;
            dto.rejectionReason = req.Rejected_Reason__c;
            dtos.add(dto);
        }

        return dtos;
    }

    private static List<Engineer_Transfer_Request__c> getSubmittedTransferRequests(
        Id requesterUserId
    ) {
        if (requesterUserId == null) {
            return new List<Engineer_Transfer_Request__c>();
        }

        return [
            SELECT Id,
                Work_Order_to_Transfer__c,
                Work_Order_to_Transfer__r.WorkOrderNumber,
                Work_Order_to_Transfer__r.Subject,
                Work_Order_to_Transfer__r.Account.Name,
                Work_Order_to_Transfer__r.Street,
                Work_Order_to_Transfer__r.City,
                Work_Order_to_Transfer__r.State,
                Work_Order_to_Transfer__r.PostalCode,
                Work_Order_to_Transfer__r.Country,
                Transfer_Work_Order_To_Engineer__r.Name,
                Accepted_on__c,
                Date_Rejected_On__c,
                Rejected_Reason__c,
                CreatedDate,
                CreatedBy.Name
            FROM Engineer_Transfer_Request__c
            WHERE CreatedById = :requesterUserId
            ORDER BY CreatedDate DESC
        ];
    }

    private static List<TransferRequestDTO> buildSubmittedTransferRequests(
        Id requesterUserId
    ) {
        List<Engineer_Transfer_Request__c> requests =
            getSubmittedTransferRequests(requesterUserId);
        return buildTransferRequests(requests);
    }

    private static Set<Id> extractWorkOrderIds(
        List<Engineer_Transfer_Request__c> requests
    ) {
        Set<Id> woIds = new Set<Id>();

        for (Engineer_Transfer_Request__c req : requests) {
            if (req.Work_Order_to_Transfer__c != null) {
                woIds.add(req.Work_Order_to_Transfer__c);
            }
        }

        return woIds;
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
        Set<Id> crewIds,
        Id ownerUserId,
        Set<Id> pendingTransferWorkOrderIds,
        Set<Id> workOrdersWithAttachment,
        Map<Id, String> quoteDownloadByWo,
        Map<Id, Id> quoteDocIdByWo,
        Map<Id, String> opportunityRecordTypeByWo
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
                Opportunity__c,
                Opportunity__r.RecordType.Name,
                Street,
                City,
                State,
                PostalCode,
                Country
            FROM WorkOrder
            WHERE OwnerId = :ownerUserId
            AND Status NOT IN :closedStatuses
            AND Id NOT IN :workOrdersWithThisTech
            LIMIT 200
        ];

        if (candidates.isEmpty()) {
            return dtos;
        }

        // Collect candidate Work Order Ids
        Set<Id> candidateWoIds = new Set<Id>();
        Set<Id> candidateOppIds = new Set<Id>();
        Map<Id, List<Id>> candidateWoByOpp = new Map<Id, List<Id>>();
        for (WorkOrder wo : candidates) {
            if (pendingTransferWorkOrderIds != null
                && pendingTransferWorkOrderIds.contains(wo.Id)) {
                // Pending transfer to this technician; keep it out of the unscheduled list
                // until the request is accepted or rejected.
                continue;
            }
            candidateWoIds.add(wo.Id);

            if (wo.Opportunity__c != null) {
                candidateOppIds.add(wo.Opportunity__c);

                List<Id> related = candidateWoByOpp.containsKey(wo.Opportunity__c)
                    ? candidateWoByOpp.get(wo.Opportunity__c)
                    : new List<Id>();
                if (!related.contains(wo.Id)) {
                    related.add(wo.Id);
                }
                candidateWoByOpp.put(wo.Opportunity__c, related);

                if (wo.Opportunity__r != null && wo.Opportunity__r.RecordType != null) {
                    opportunityRecordTypeByWo.put(
                        wo.Id,
                        wo.Opportunity__r.RecordType.Name
                    );
                }
            }
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

        Map<Id, Boolean> saHasSchedule = new Map<Id, Boolean>();

        List<ServiceAppointment> candidateSAs = [
            SELECT Id,
                Work_Order__c,
                Status,
                SchedStartTime,
                SchedEndTime
            FROM ServiceAppointment
            WHERE Work_Order__c IN :candidateWoIds
            AND Status NOT IN :saClosedStatuses
        ];

        for (ServiceAppointment sa : candidateSAs) {
            if (sa.Work_Order__c != null) {
                saIdToWoId.put(sa.Id, sa.Work_Order__c);
            }

            Boolean hasSchedule =
                (sa.SchedStartTime != null) || (sa.SchedEndTime != null);
            saHasSchedule.put(sa.Id, hasSchedule);
        }

        if (!candidateWoIds.isEmpty()) {
            for (ContentDocumentLink cdl : [
                SELECT LinkedEntityId,
                    ContentDocumentId,
                    ContentDocument.LatestPublishedVersionId
                FROM ContentDocumentLink
                WHERE LinkedEntityId IN :candidateWoIds
            ]) {
                if (cdl.LinkedEntityId == null) continue;

                workOrdersWithAttachment.add(cdl.LinkedEntityId);

                if (cdl.ContentDocument != null &&
                    String.isNotBlank(cdl.ContentDocument.LatestPublishedVersionId) &&
                    !quoteDownloadByWo.containsKey(cdl.LinkedEntityId)) {

                    quoteDownloadByWo.put(
                        cdl.LinkedEntityId,
                        '/sfc/servlet.shepherd/document/download/' +
                            cdl.ContentDocumentId
                    );

                    if (cdl.ContentDocumentId != null) {
                        quoteDocIdByWo.put(
                            cdl.LinkedEntityId,
                            cdl.ContentDocumentId
                        );
                    }
                }
            }
        }

        if (!candidateOppIds.isEmpty()) {
            for (ContentDocumentLink cdl : [
                SELECT LinkedEntityId,
                    ContentDocumentId,
                    ContentDocument.LatestPublishedVersionId
                FROM ContentDocumentLink
                WHERE LinkedEntityId IN :candidateOppIds
            ]) {
                if (cdl.LinkedEntityId == null ||
                    !candidateWoByOpp.containsKey(cdl.LinkedEntityId)) {
                    continue;
                }

                for (Id woId : candidateWoByOpp.get(cdl.LinkedEntityId)) {
                    workOrdersWithAttachment.add(woId);

                    if (cdl.ContentDocument != null &&
                        String.isNotBlank(cdl.ContentDocument.LatestPublishedVersionId) &&
                        !quoteDownloadByWo.containsKey(woId)) {

                        quoteDownloadByWo.put(
                            woId,
                            '/sfc/servlet.shepherd/document/download/' +
                                cdl.ContentDocumentId
                        );

                        if (cdl.ContentDocumentId != null) {
                            quoteDocIdByWo.put(
                                woId,
                                cdl.ContentDocumentId
                            );
                        }
                    }
                }
            }
        }

        // 3) For those ServiceAppointments, see which have any AssignedResource rows
        Set<Id> workOrdersWithNonCrewAssignment = new Set<Id>();
        Set<Id> workOrdersWithCrewAssignment     = new Set<Id>();

        if (!saIdToWoId.isEmpty()) {
            for (AssignedResource ar : [
                SELECT ServiceAppointmentId,
                    ServiceCrewId,
                    ServiceResource.ResourceType,
                    ServiceResource.RelatedRecordId
                FROM AssignedResource
                WHERE ServiceAppointmentId IN :saIdToWoId.keySet()
            ]) {
                Id woId = saIdToWoId.get(ar.ServiceAppointmentId);
                if (woId != null) {
                    Boolean isCrewLike =
                        ar.ServiceCrewId != null ||
                        (ar.ServiceResource != null &&
                            (ar.ServiceResource.ResourceType == 'Crew' ||
                             ar.ServiceResource.ResourceType == 'ServiceCrew' ||
                             String.valueOf(ar.ServiceResource.RelatedRecordId).startsWith('00G')));

                    if (isCrewLike) {
                        workOrdersWithCrewAssignment.add(woId);
                    } else {
                        Boolean hasSchedule = saHasSchedule.containsKey(ar.ServiceAppointmentId)
                            ? saHasSchedule.get(ar.ServiceAppointmentId)
                            : false;

                        if (hasSchedule) {
                            workOrdersWithNonCrewAssignment.add(woId);
                        }
                    }
                }
            }
        }

        // 4) Build DTOs only for Work Orders that:
        //    - are owned by the user
        //    - are not closed
        //    - are not already on this tech calendar (workOrdersWithThisTech)
        //    - do not have any open ServiceAppointment with an assigned individual resource
        //      (but allow ones already placed with a crew/queue so they can be picked up)
        for (WorkOrder wo : candidates) {
            if (workOrdersWithNonCrewAssignment.contains(wo.Id)) {
                // Already has an open Service Appointment with assigned resources
                // so it should not appear as needs scheduling
                continue;
            }

            UnscheduledWorkOrderDTO dto = new UnscheduledWorkOrderDTO();
            dto.workOrderId     = wo.Id;
            dto.workOrderNumber = wo.WorkOrderNumber;
            dto.subject         = wo.Subject;
            dto.status          = wo.Status;
            dto.hasCrewAssignment = workOrdersWithCrewAssignment.contains(wo.Id);
            dto.accountId       = wo.AccountId;
            dto.accountName     = (wo.Account != null ? wo.Account.Name : null);
            dto.opportunityRecordType = (wo.Opportunity__c != null &&
                opportunityRecordTypeByWo.containsKey(wo.Id))
                    ? opportunityRecordTypeByWo.get(wo.Id)
                    : null;
            dto.hasQuoteAttachment =
                workOrdersWithAttachment != null &&
                workOrdersWithAttachment.contains(wo.Id);
            dto.quoteAttachmentDownloadUrl = (wo.Id != null &&
                quoteDownloadByWo != null && quoteDownloadByWo.containsKey(wo.Id))
                    ? quoteDownloadByWo.get(wo.Id)
                    : null;
            dto.quoteAttachmentDocumentId = (wo.Id != null &&
                quoteDocIdByWo != null && quoteDocIdByWo.containsKey(wo.Id))
                    ? quoteDocIdByWo.get(wo.Id)
                    : null;
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
        Datetime endDateTimeIso,
        Id targetUserId
    ) {
        if (workOrderId == null || startDateTimeIso == null) {
            throw new AuraHandledException('Work Order and start date/time are required.');
        }

        Id owningUserId = (targetUserId != null)
            ? targetUserId
            : UserInfo.getUserId();

        // Current user's ServiceResource
        ServiceResource sr = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId = :owningUserId
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

        Integer durationMs = 6 * 60 * 60 * 1000; // 6 hour default
        if (endDateTimeIso != null && endDateTimeIso.getTime() > startDateTimeIso.getTime()) {
            durationMs = Integer.valueOf(endDateTimeIso.getTime() - startDateTimeIso.getTime());
        }

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c  = wo.Id,
            Subject        = wo.Subject,
            Street         = wo.Street,
            City           = wo.City,
            State          = wo.State,
            PostalCode     = wo.PostalCode,
            Country        = wo.Country,
            ParentRecordId = wo.Id,
            Status         = 'New'
        );
        insert sa;

        Datetime endDateTime = startDateTimeIso.addSeconds(durationMs / 1000);

        ServiceAppointment saTimeUpdate = new ServiceAppointment(
            Id              = sa.Id,
            SchedStartTime  = startDateTimeIso,
            SchedEndTime    = endDateTime
        );
        update saTimeUpdate;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = sr.Id
        );
        insert ar;

        ServiceAppointment saStatusUpdate = new ServiceAppointment(
            Id      = sa.Id,
            Status  = 'Scheduled'
        );
        update saStatusUpdate;

        return sa.Id;
    }

    @AuraEnabled
    public static void unassignAppointment(Id appointmentId) {
        if (appointmentId == null) {
            throw new AuraHandledException('Appointment ID is required.');
        }

        ServiceAppointment sa = [
            SELECT Id
            FROM ServiceAppointment
            WHERE Id = :appointmentId
            LIMIT 1
            FOR UPDATE
        ];

        List<AssignedResource> ars = [
            SELECT Id
            FROM AssignedResource
            WHERE ServiceAppointmentId = :appointmentId
            FOR UPDATE
        ];

        if (!ars.isEmpty()) {
            delete ars;
        }

        sa = new ServiceAppointment(
            Id = appointmentId,
            Status = 'New',
            SchedStartTime = null,
            SchedEndTime = null
        );

        update sa;
    }

    @AuraEnabled
    public static void updateResourceAbsence(
        Id absenceId,
        Datetime startDateTimeIso,
        Datetime endDateTimeIso
    ) {
        if (absenceId == null) {
            throw new AuraHandledException('Absence ID is required.');
        }

        if (startDateTimeIso == null || endDateTimeIso == null) {
            throw new AuraHandledException('Start and End date/time are required.');
        }

        if (endDateTimeIso.getTime() <= startDateTimeIso.getTime()) {
            throw new AuraHandledException('End time must be after start time.');
        }

        ResourceAbsence ra = [
            SELECT Id
            FROM ResourceAbsence
            WHERE Id = :absenceId
            LIMIT 1
            FOR UPDATE
        ];

        ra = new ResourceAbsence(
            Id = ra.Id,
            Start = startDateTimeIso,
            End = endDateTimeIso
        );

        update ra;
    }

    @AuraEnabled
    public static void deleteResourceAbsence(Id absenceId) {
        if (absenceId == null) {
            throw new AuraHandledException('Absence ID is required.');
        }

        ResourceAbsence ra = [
            SELECT Id
            FROM ResourceAbsence
            WHERE Id = :absenceId
            LIMIT 1
            FOR UPDATE
        ];

        delete ra;
    }

    @AuraEnabled
    public static List<ServiceResourceOption> getTerritoryResources(Id workOrderId) {
        if (workOrderId == null) {
            throw new AuraHandledException('Work Order Id is required.');
        }

        WorkOrder wo = [
            SELECT Id, ServiceTerritoryId
            FROM WorkOrder
            WHERE Id = :workOrderId
            LIMIT 1
        ];

        if (wo.ServiceTerritoryId == null) {
            return new List<ServiceResourceOption>();
        }

        List<ServiceResourceOption> options = new List<ServiceResourceOption>();
        for (ServiceTerritoryMember stm : [
            SELECT Id, ServiceResourceId, ServiceResource.Name
                , ServiceResource.RelatedRecordId
            FROM ServiceTerritoryMember
            WHERE ServiceTerritoryId = :wo.ServiceTerritoryId
            AND EffectiveStartDate != null
            AND EffectiveStartDate <= :Date.today()
            AND (EffectiveEndDate = null OR EffectiveEndDate >= :Date.today())
            ORDER BY ServiceResource.Name
        ]) {
            if (stm.ServiceResourceId == null) continue;

            // Only include resources that are linked to a User (not crews/queues)
            Id relatedId = (stm.ServiceResource != null)
                ? stm.ServiceResource.RelatedRecordId
                : null;

            if (relatedId == null || relatedId.getSObjectType() != User.sObjectType) {
                continue;
            }

            ServiceResourceOption opt = new ServiceResourceOption();
            opt.serviceResourceId = stm.ServiceResourceId;
            opt.name = (stm.ServiceResource != null) ? stm.ServiceResource.Name : null;
            opt.userId = relatedId;
            options.add(opt);
        }

        return options;
    }

    @AuraEnabled
    public static void requestRescheduleToResource(
        Id workOrderId,
        Id serviceResourceId
    ) {
        if (workOrderId == null || serviceResourceId == null) {
            throw new AuraHandledException(
                'Work Order and Service Resource are required to request a reschedule.'
            );
        }

        ServiceResource sr = [
            SELECT Id, RelatedRecordId
            FROM ServiceResource
            WHERE Id = :serviceResourceId
            LIMIT 1
        ];

        WorkOrder wo = [
            SELECT Id, OwnerId
            FROM WorkOrder
            WHERE Id = :workOrderId
            LIMIT 1
            FOR UPDATE
        ];

        if (sr.RelatedRecordId != null) {
            wo.OwnerId = sr.RelatedRecordId;
            update wo;
        }

        Set<String> closedStatuses = new Set<String>{
            'Completed',
            'Completed WO',
            'Closed',
            'Cannot Complete',
            'Canceled',
            'Closed - Submitted'
        };

        List<ServiceAppointment> openAppointments = [
            SELECT Id
            FROM ServiceAppointment
            WHERE Work_Order__c = :workOrderId
            AND Status NOT IN :closedStatuses
            FOR UPDATE
        ];

        if (openAppointments.isEmpty()) {
            return;
        }

        Set<Id> openApptIds = new Set<Id>();
        for (ServiceAppointment sa : openAppointments) {
            openApptIds.add(sa.Id);
        }

        List<AssignedResource> ars = [
            SELECT Id
            FROM AssignedResource
            WHERE ServiceAppointmentId IN :openApptIds
        ];

        if (!ars.isEmpty()) {
            delete ars;
        }

        for (ServiceAppointment sa : openAppointments) {
            sa.Status = 'Canceled';
            sa.SchedStartTime = null;
            sa.SchedEndTime = null;
        }

        update openAppointments;
    }

    @AuraEnabled
    public static void markWorkOrderQuoteSent(Id workOrderId) {
        if (workOrderId == null) {
            throw new AuraHandledException('Work Order Id is required.');
        }

        WorkOrder wo = [
            SELECT Id, Status
            FROM WorkOrder
            WHERE Id = :workOrderId
            LIMIT 1
            FOR UPDATE
        ];

        wo.Status = 'Quote Sent';
        update wo;
    }

    @AuraEnabled
    public static void createEngineerTransferRequest(
        Id workOrderId,
        Id targetUserId
    ) {
        if (workOrderId == null || targetUserId == null) {
            throw new AuraHandledException(
                'Work Order and target technician are required to submit a transfer request.'
            );
        }

        List<Engineer_Transfer_Request__c> existing = [
            SELECT Id
            FROM Engineer_Transfer_Request__c
            WHERE Work_Order_to_Transfer__c = :workOrderId
            AND Transfer_Work_Order_To_Engineer__c = :targetUserId
            AND Accepted_on__c = null
            AND Date_Rejected_On__c = null
            LIMIT 1
        ];

        if (!existing.isEmpty()) {
            return;
        }

        Engineer_Transfer_Request__c req = new Engineer_Transfer_Request__c();
        req.Work_Order_to_Transfer__c = workOrderId;
        req.Transfer_Work_Order_To_Engineer__c = targetUserId;

        insert req;
    }

    @AuraEnabled
    public static void acceptEngineerTransferRequest(
        Id transferRequestId,
        Id targetOwnerId
    ) {
        if (transferRequestId == null) {
            throw new AuraHandledException('Transfer Request Id is required to accept.');
        }

        Engineer_Transfer_Request__c req = [
            SELECT Id,
                Work_Order_to_Transfer__c,
                Transfer_Work_Order_To_Engineer__c
            FROM Engineer_Transfer_Request__c
            WHERE Id = :transferRequestId
            LIMIT 1
            FOR UPDATE
        ];

        Id currentUserId = UserInfo.getUserId();
        Id targetUserId = req.Transfer_Work_Order_To_Engineer__c;

        if (currentUserId != targetUserId) {
            User targetUser = targetUserId != null
                ? [SELECT Id, ManagerId FROM User WHERE Id = :targetUserId LIMIT 1]
                : null;

            Boolean isManagerActingForTarget =
                targetUser != null && targetUser.ManagerId == currentUserId;

            if (!isManagerActingForTarget) {
                throw new AuraHandledException(
                    'You are not authorized to accept this request.'
                );
            }
        }

        req.Accepted_on__c = Datetime.now();
        req.Date_Rejected_On__c = null;
        req.Rejected_Reason__c = null;
        update req;

        if (req.Work_Order_to_Transfer__c == null) {
            return;
        }

        WorkOrder wo = [
            SELECT Id
            FROM WorkOrder
            WHERE Id = :req.Work_Order_to_Transfer__c
            LIMIT 1
            FOR UPDATE
        ];

        Id effectiveOwnerId = req.Transfer_Work_Order_To_Engineer__c != null
            ? req.Transfer_Work_Order_To_Engineer__c
            : currentUserId;

        if (targetOwnerId != null && targetOwnerId == req.Transfer_Work_Order_To_Engineer__c) {
            effectiveOwnerId = targetOwnerId;
        }

        wo.OwnerId = effectiveOwnerId;
        update wo;

        Set<String> closedStatuses = new Set<String>{
            'Completed',
            'Completed WO',
            'Closed',
            'Cannot Complete',
            'Canceled',
            'Closed - Submitted'
        };

        List<ServiceAppointment> openAppointments = [
            SELECT Id
            FROM ServiceAppointment
            WHERE Work_Order__c = :req.Work_Order_to_Transfer__c
            AND Status NOT IN :closedStatuses
            FOR UPDATE
        ];

        if (openAppointments.isEmpty()) {
            return;
        }

        Set<Id> openApptIds = new Set<Id>();
        for (ServiceAppointment sa : openAppointments) {
            openApptIds.add(sa.Id);
        }

        List<AssignedResource> ars = [
            SELECT Id
            FROM AssignedResource
            WHERE ServiceAppointmentId IN :openApptIds
        ];

        if (!ars.isEmpty()) {
            delete ars;
        }

        for (ServiceAppointment sa : openAppointments) {
            sa.Status = 'Canceled';
            sa.SchedStartTime = null;
            sa.SchedEndTime = null;
        }

        update openAppointments;
    }

    @AuraEnabled
    public static void rejectEngineerTransferRequest(
        Id transferRequestId,
        String reason
    ) {
        if (transferRequestId == null) {
            throw new AuraHandledException('Transfer Request Id is required to reject.');
        }

        if (String.isBlank(reason)) {
            throw new AuraHandledException('A rejection reason is required.');
        }

        Engineer_Transfer_Request__c req = [
            SELECT Id,
                Transfer_Work_Order_To_Engineer__c
            FROM Engineer_Transfer_Request__c
            WHERE Id = :transferRequestId
            LIMIT 1
            FOR UPDATE
        ];

        Id currentUserId = UserInfo.getUserId();
        Id targetUserId = req.Transfer_Work_Order_To_Engineer__c;

        if (currentUserId != targetUserId) {
            User targetUser = targetUserId != null
                ? [SELECT Id, ManagerId FROM User WHERE Id = :targetUserId LIMIT 1]
                : null;

            Boolean isManagerActingForTarget =
                targetUser != null && targetUser.ManagerId == currentUserId;

            if (!isManagerActingForTarget) {
                throw new AuraHandledException(
                    'You are not authorized to reject this request.'
                );
            }
        }

        req.Date_Rejected_On__c = Datetime.now();
        req.Rejected_Reason__c = reason;
        req.Accepted_on__c = null;

        update req;
    }
}
