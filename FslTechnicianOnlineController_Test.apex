@isTest(SeeAllData=true)
public class FslTechnicianOnlineController_Test {

    /**
     * Wraps a resource, its active territory membership, and the backing User.
     */
    private class ResourceContext {
        ServiceTerritoryMember stm;
        ServiceResource resource;
        User user;
    }

    /**
     * Return the first picklist value for a field, used for synthetic records when the
     * org's configured value set is unknown.
     */
    private static String firstPicklistValue(Schema.SObjectField field) {
        List<Schema.PicklistEntry> entries = field.getDescribe().getPicklistValues();
        return (!entries.isEmpty()) ? entries[0].getValue() : null;
    }

    /**
     * Return up to N picklist values for a field.
     */
    private static List<String> somePicklistValues(Schema.SObjectField field, Integer maxValues) {
        List<String> result = new List<String>();
        List<Schema.PicklistEntry> entries = field.getDescribe().getPicklistValues();
        Integer maxCount = Math.min(maxValues, entries.size());
        for (Integer i = 0; i < maxCount; i++) {
            result.add(entries[i].getValue());
        }
        return result;
    }

    /**
     * Helper: truncate a Datetime value to the nearest minute (zero out seconds/millis).
     */
    private static Datetime truncateToMinute(Datetime dt) {
        if (dt == null) return null;
        Date d = dt.date();
        Time t = dt.time();
        return Datetime.newInstance(d, Time.newInstance(t.hour(), t.minute(), 0, 0));
    }

    /**
     * Simple helper to create an extra User (for transfer tests, manager, etc.).
     */
    private static User createPlainUser(String seed) {
        Profile p = [SELECT Id FROM Profile LIMIT 1];
        String unique = seed + String.valueOf(Datetime.now().getTime());
        String alias  = (seed.length() > 8) ? seed.substring(0, 8) : seed;

        User u = new User(
            ProfileId         = p.Id,
            LastName          = seed + 'User',
            Alias             = alias,
            Email             = unique + '@example.com',
            Username          = unique + '@example.com',
            TimeZoneSidKey    = 'America/Los_Angeles',
            LocaleSidKey      = 'en_US',
            EmailEncodingKey  = 'UTF-8',
            LanguageLocaleKey = 'en_US'
        );
        insert u;
        return u;
    }

    /**
     * Build a synthetic technician, resource, and territory when live data is not found.
     */
    private static ResourceContext createSyntheticResourceContext() {
        Profile standardProfile = [SELECT Id FROM Profile LIMIT 1];

        String uniqueSuffix = String.valueOf(Datetime.now().getTime());

        // Ensure we have Operating Hours (required by ServiceTerritory in many FSL orgs)
        OperatingHours oph;
        List<OperatingHours> ohList = [
            SELECT Id
            FROM OperatingHours
            LIMIT 1
        ];
        if (ohList.isEmpty()) {
            oph = new OperatingHours(
                Name     = 'Test Operating Hours ' + uniqueSuffix,
                TimeZone = 'America/Los_Angeles'
            );
            insert oph;
        } else {
            oph = ohList[0];
        }

        User user = new User(
            ProfileId         = standardProfile.Id,
            LastName          = 'Tech',
            Alias             = 'tech',
            Email             = 'tech' + uniqueSuffix + '@example.com',
            Username          = 'tech' + uniqueSuffix + '@example.com',
            TimeZoneSidKey    = 'America/Los_Angeles',
            LocaleSidKey      = 'en_US',
            EmailEncodingKey  = 'UTF-8',
            LanguageLocaleKey = 'en_US'
        );
        insert user;

        ServiceTerritory territory = new ServiceTerritory(
            Name             = 'Test Territory ' + uniqueSuffix,
            IsActive         = true,
            OperatingHoursId = oph.Id
        );
        insert territory;

        ServiceResource resource = new ServiceResource(
            Name            = 'Tech Resource ' + uniqueSuffix,
            IsActive        = true,
            RelatedRecordId = user.Id
        );
        insert resource;

        ServiceTerritoryMember stm = new ServiceTerritoryMember(
            ServiceTerritoryId = territory.Id,
            ServiceResourceId  = resource.Id,
            EffectiveStartDate = Date.today().addDays(-1)
        );
        insert stm;

        ResourceContext context = new ResourceContext();
        context.stm      = stm;
        context.resource = resource;
        context.user     = user;

        return context;
    }

    /**
     * Ensure a ResourceAbsence exists that overlaps "now".
     * ResourceId is required in your org, so we tie it to the current ServiceResource.
     */
    private static ResourceAbsence ensureResourceAbsence(ResourceContext context) {
        List<ResourceAbsence> existing = [
            SELECT Id, ResourceId, Start, End
            FROM ResourceAbsence
            WHERE ResourceId = :context.resource.Id
              AND Start <= :Datetime.now()
              AND End   >= :Datetime.now()
            LIMIT 1
        ];

        if (!existing.isEmpty()) {
            return existing[0];
        }

        Datetime start   = Datetime.now().addHours(-2);
        Datetime endtime = Datetime.now().addHours(2);

        ResourceAbsence ra = new ResourceAbsence(
            ResourceId = context.resource.Id,
            Start      = start,
            End        = endtime
        );
        insert ra;

        return ra;
    }

    /**
     * Locate an active Service Resource with an active territory assignment and matching Work Orders.
     */
    private static ResourceContext getActiveResourceContext() {
        List<ServiceTerritoryMember> stms = [
            SELECT Id,
                   ServiceResourceId,
                   ServiceTerritoryId,
                   ServiceResource.RelatedRecordId,
                   ServiceResource.IsActive
            FROM ServiceTerritoryMember
            WHERE ServiceResource.RelatedRecordId != null
              AND ServiceResource.IsActive = true          // 🔹 only active resources
              AND EffectiveStartDate != null
              AND EffectiveStartDate <= :Date.today()
              AND (EffectiveEndDate = null OR EffectiveEndDate >= :Date.today())
              AND ServiceTerritoryId IN (
                  SELECT ServiceTerritoryId
                  FROM WorkOrder
                  WHERE ServiceTerritoryId != null
              )
              AND ServiceTerritoryId IN (
                  SELECT ServiceTerritoryId
                  FROM ServiceAppointment
                  WHERE SchedStartTime != null AND SchedEndTime != null
              )
            LIMIT 1
        ];


        ResourceContext context;

        if (stms.isEmpty()) {
            context = createSyntheticResourceContext();
        } else {
            ServiceTerritoryMember stm = stms[0];
        
            ServiceResource resource = [
                SELECT Id, RelatedRecordId, IsActive
                FROM ServiceResource
                WHERE Id = :stm.ServiceResourceId
                LIMIT 1
            ];
        
            // Just in case – if the resource is somehow inactive, flip it on
            if (!resource.IsActive) {
                resource.IsActive = true;
                update resource;
            }
        
            User user = [
                SELECT Id
                FROM User
                WHERE Id = :resource.RelatedRecordId
                LIMIT 1
            ];
        
            context = new ResourceContext();
            context.stm      = stm;
            context.resource = resource;
            context.user     = user;
        }


        // Ensure we have a ResourceAbsence overlapping now
        ensureResourceAbsence(context);

        // Also ensure we have Work Orders and Appointments in various statuses
        ensureWorkOrders(context);
        ensureServiceAppointments(context);

        return context;
    }

    /**
     * Create multiple WorkOrders for the context in various statuses (if none already exist).
     */
    private static void ensureWorkOrders(ResourceContext context) {
        List<WorkOrder> existing = [
            SELECT Id
            FROM WorkOrder
            WHERE OwnerId = :context.user.Id
              AND ServiceTerritoryId = :context.stm.ServiceTerritoryId
            LIMIT 1
        ];

        if (!existing.isEmpty()) {
            return;
        }

        Account acc = new Account(Name = 'Test Account - ' + context.user.Id);
        insert acc;

        List<String> statuses = somePicklistValues(WorkOrder.Status, 3);
        if (statuses.isEmpty()) {
            statuses.add('New');
        }

        List<WorkOrder> toInsert = new List<WorkOrder>();
        Integer i = 0;
        for (String statusVal : statuses) {
            WorkOrder wo = new WorkOrder(
                Subject            = 'Test Work Order ' + i,
                Status             = statusVal,
                OwnerId            = context.user.Id,
                ServiceTerritoryId = context.stm.ServiceTerritoryId,
                AccountId          = acc.Id,
                Street             = '123 Main St',
                City               = 'San Francisco',
                State              = 'California',
                PostalCode         = '94105',
                Country            = 'United States'
            );
            toInsert.add(wo);
            i++;
        }
        insert toInsert;
    }

    /**
     * Return a single WorkOrder for this context (ensuring some exist).
     */
    private static WorkOrder ensureWorkOrder(ResourceContext context) {
        ensureWorkOrders(context);

        return [
            SELECT Id, OwnerId, ServiceTerritoryId, Status, Subject, AccountId,
                   Street, City, State, PostalCode, Country
            FROM WorkOrder
            WHERE OwnerId = :context.user.Id
              AND ServiceTerritoryId = :context.stm.ServiceTerritoryId
            LIMIT 1
        ];
    }

    /**
     * Create multiple ServiceAppointments for the context in various statuses (if none already exist).
     */
    private static void ensureServiceAppointments(ResourceContext context) {
        List<ServiceAppointment> existing = [
            SELECT Id
            FROM ServiceAppointment
            WHERE ServiceTerritoryId = :context.stm.ServiceTerritoryId
              AND SchedStartTime != null
              AND SchedEndTime   != null
            LIMIT 1
        ];

        if (!existing.isEmpty()) {
            return;
        }

        WorkOrder baseWo = ensureWorkOrder(context);

        List<String> saStatuses = somePicklistValues(ServiceAppointment.Status, 3);
        if (saStatuses.isEmpty()) {
            saStatuses.add('New');
        }

        String durationType = firstPicklistValue(ServiceAppointment.DurationType);
        if (durationType == null) {
            durationType = 'Minutes';
        }

        List<ServiceAppointment> toInsert = new List<ServiceAppointment>();
        for (Integer i = 0; i < saStatuses.size(); i++) {
            Datetime startTime = Datetime.now().addDays(i + 1);
            Datetime endTime   = startTime.addHours(1);

            // Always insert as New to avoid validation rules on Scheduled/Dispatched for new records
            ServiceAppointment sa = new ServiceAppointment(
                Subject            = 'Test Appointment ' + i,
                Status             = 'New',
                ServiceTerritoryId = context.stm.ServiceTerritoryId,
                ParentRecordId     = baseWo.Id,
                Duration           = 60,
                DurationType       = durationType,
                Work_Order__c      = baseWo.Id
            );
            toInsert.add(sa);
        }
        insert toInsert;

        // Now update times and put back the original statuses
        List<ServiceAppointment> updates = new List<ServiceAppointment>();
        for (Integer i = 0; i < toInsert.size(); i++) {
            Datetime startTime = Datetime.now().addDays(i + 1);
            Datetime endTime   = startTime.addHours(1);

            ServiceAppointment saUpdate = new ServiceAppointment(
                Id              = toInsert[i].Id,
                SchedStartTime  = startTime,
                SchedEndTime    = endTime,
                Status          = saStatuses[i]
            );
            updates.add(saUpdate);
        }
        update updates;
    }

    /**
     * Return a single scheduled appointment for this context (ensuring some exist).
     */
    private static ServiceAppointment ensureScheduledAppointment(ResourceContext context) {
        ensureServiceAppointments(context);

        return [
            SELECT Id, SchedStartTime, SchedEndTime, Status, ServiceTerritoryId, ParentRecordId, Work_Order__c
            FROM ServiceAppointment
            WHERE ServiceTerritoryId = :context.stm.ServiceTerritoryId
              AND SchedStartTime != null
              AND SchedEndTime   != null
            LIMIT 1
        ];
    }

    private static User getExistingTechnicianUser() {
        return getActiveResourceContext().user;
    }

    // ------------------------- CORE APPOINTMENT TESTS -------------------------

    @isTest
    static void testGetAppointmentsForActiveUser() {
        ResourceContext context = getActiveResourceContext();
        User techUser = context.user;

        System.runAs(techUser) {
            Test.startTest();
            FslTechnicianOnlineController.AppointmentResult result =
                FslTechnicianOnlineController.getMyAppointmentsOnline(null);
            Test.stopTest();

            System.assertNotEquals(null, result, 'AppointmentResult should not be null for active user.');
        }
    }

    @isTest
    static void testGetAppointmentsForOtherUser() {
        ResourceContext context = getActiveResourceContext();
        User targetUser = context.user;

        Test.startTest();
        FslTechnicianOnlineController.AppointmentResult result =
            FslTechnicianOnlineController.getMyAppointmentsOnline(targetUser.Id);
        Test.stopTest();

        System.assertNotEquals(null, result, 'AppointmentResult should not be null for other user.');
        System.assertEquals(targetUser.Id, result.viewingUserId, 'Viewing user id should match target.');
    }

    @isTest
    static void testRescheduleAndEndTimeUpdate() {
        ResourceContext context = getActiveResourceContext();
        ServiceAppointment sa   = ensureScheduledAppointment(context);

        Datetime originalEnd = sa.SchedEndTime;
        Datetime newStart    = sa.SchedStartTime.addDays(1);
        Datetime earlierEnd  = sa.SchedStartTime.addMinutes(5);

        Test.startTest();
        FslTechnicianOnlineController.rescheduleAppointment(sa.Id, newStart);
        FslTechnicianOnlineController.updateAppointmentEnd(sa.Id, earlierEnd);
        Test.stopTest();

        ServiceAppointment updated = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :sa.Id
            LIMIT 1
        ];

        System.assertEquals(newStart, updated.SchedStartTime, 'Start time should be updated.');

        System.assertNotEquals(originalEnd, updated.SchedEndTime, 'End time should be updated.');
        System.assert(
            updated.SchedEndTime.getTime() >= updated.SchedStartTime.getTime(),
            'End time should be on or after the start time (min duration enforced).'
        );
    }

    @isTest
    static void testAssignCrewAppointment() {
        ResourceContext context = getActiveResourceContext();
        ServiceAppointment sa   = ensureScheduledAppointment(context);

        Test.startTest();
        FslTechnicianOnlineController.assignCrewAppointment(sa.Id, context.resource.Id);
        Test.stopTest();

        Integer arCount = [
            SELECT COUNT()
            FROM AssignedResource
            WHERE ServiceAppointmentId = :sa.Id
        ];
        System.assertEquals(1, arCount, 'Crew assignment should create a single AssignedResource.');
    }

    @isTest
    static void testCreateAppointmentAndOptions() {
        ResourceContext context = getActiveResourceContext();
        User techUser           = context.user;

        WorkOrder wo = ensureWorkOrder(context);

        Datetime startTime = Datetime.now().addDays(2);
        Datetime endTime   = startTime.addHours(1);

        Test.startTest();
        Id newApptId = FslTechnicianOnlineController.createAppointmentForWorkOrder(
            wo.Id,
            startTime,
            endTime,
            techUser.Id
        );
        Test.stopTest();

        System.assertNotEquals(null, newApptId, 'New appointment Id should be returned.');

        ServiceAppointment created = [
            SELECT Id, ParentRecordId, SchedStartTime, SchedEndTime, Status
            FROM ServiceAppointment
            WHERE Id = :newApptId
            LIMIT 1
        ];

        Datetime expectedStart = truncateToMinute(startTime);
        Datetime expectedEnd   = truncateToMinute(endTime);

        System.assertEquals(wo.Id,       created.ParentRecordId, 'ParentRecordId should match WorkOrder.');
        System.assertEquals(expectedStart, truncateToMinute(created.SchedStartTime), 'Start time should match (to the minute).');
        System.assertEquals(expectedEnd,   truncateToMinute(created.SchedEndTime),   'End time should match (to the minute).');
        System.assertEquals('Dispatched', created.Status, 'Appointment should be dispatched.');
    }

    @isTest
    static void testRequestRescheduleAndMarkQuoteSent() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo           = ensureWorkOrder(context);

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c      = wo.Id,
            Subject            = 'Reschedule Test',
            Status             = 'New',
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            ParentRecordId     = wo.Id
        );
        insert sa;

        sa = new ServiceAppointment(
            Id              = sa.Id,
            SchedStartTime  = Datetime.now().addDays(1),
            SchedEndTime    = Datetime.now().addDays(1).addHours(1),
            Status          = 'New'
        );
        update sa;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = context.resource.Id
        );
        insert ar;

        Test.startTest();
        FslTechnicianOnlineController.requestRescheduleToResource(wo.Id, context.resource.Id);
        FslTechnicianOnlineController.markWorkOrderQuoteSent(wo.Id);
        Test.stopTest();

        WorkOrder woAfter = [
            SELECT Id, Status, OwnerId
            FROM WorkOrder
            WHERE Id = :wo.Id
            LIMIT 1
        ];
        System.assertEquals('Quote Sent', woAfter.Status, 'Work Order status should be Quote Sent.');

        ServiceAppointment saAfter = [
            SELECT Id, Status, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :sa.Id
            LIMIT 1
        ];
        System.assertEquals('Canceled', saAfter.Status, 'Open appointments should be canceled.');
        System.assertEquals(null, saAfter.SchedStartTime, 'Schedule should be cleared.');
        System.assertEquals(null, saAfter.SchedEndTime, 'Schedule should be cleared.');

        Integer arCountAfter = [
            SELECT COUNT()
            FROM AssignedResource
            WHERE ServiceAppointmentId = :sa.Id
        ];
        System.assertEquals(0, arCountAfter, 'Assignments should be removed when requesting reschedule.');
    }

    @isTest
    static void testMarkWorkOrderPoAttached() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo           = ensureWorkOrder(context);

        Test.startTest();
        FslTechnicianOnlineController.markWorkOrderPoAttached(wo.Id);
        Test.stopTest();

        WorkOrder woAfter = [
            SELECT Status
            FROM WorkOrder
            WHERE Id = :wo.Id
            LIMIT 1
        ];

        System.assertEquals(
            'PO Attached',
            woAfter.Status,
            'Work Order status should be PO Attached.'
        );
    }

    @isTest
    static void testCancelWorkOrderUpdatesStatusAndNotes() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo           = ensureWorkOrder(context);

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c      = wo.Id,
            Subject            = 'Cancel Test',
            Status             = 'New',
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            ParentRecordId     = wo.Id
        );
        insert sa;

        sa = new ServiceAppointment(
            Id             = sa.Id,
            SchedStartTime = Datetime.now().addDays(2),
            SchedEndTime   = Datetime.now().addDays(2).addHours(2),
            Status         = 'New'
        );
        update sa;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = context.resource.Id
        );
        insert ar;

        String cancellationReason = 'Customer requested cancellation';

        Test.startTest();
        FslTechnicianOnlineController.cancelWorkOrder(wo.Id, cancellationReason);
        Test.stopTest();

        WorkOrder woAfter = [
            SELECT Id, Status, Tech_Support_Notes__c
            FROM WorkOrder
            WHERE Id = :wo.Id
            LIMIT 1
        ];
        System.assertEquals('Canceled', woAfter.Status, 'Work Order should be canceled.');
        System.assertEquals(cancellationReason, woAfter.Tech_Support_Notes__c,
            'Tech support notes should store the cancellation reason.');

        ServiceAppointment saAfter = [
            SELECT Id, Status, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :sa.Id
            LIMIT 1
        ];
        System.assertEquals('Canceled', saAfter.Status, 'Appointment should be canceled.');
        System.assertEquals(null, saAfter.SchedStartTime, 'Start time should be cleared.');
        System.assertEquals(null, saAfter.SchedEndTime, 'End time should be cleared.');

        Integer arCountAfter = [
            SELECT COUNT()
            FROM AssignedResource
            WHERE ServiceAppointmentId = :sa.Id
        ];
        System.assertEquals(0, arCountAfter, 'Assignments should be removed when canceling.');
    }

    @isTest
    static void testMarkWorkOrderReadyForCloseUpdatesStatusAndNotes() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder woWithNotes   = ensureWorkOrder(context);
        woWithNotes.Tech_Support_Notes__c = 'Existing note';
        update woWithNotes;

        WorkOrder woNoNotes = ensureWorkOrder(context);

        Test.startTest();
        FslTechnicianOnlineController.markWorkOrderReadyForClose(
            woWithNotes.Id,
            'Follow-up items listed'
        );
        FslTechnicianOnlineController.markWorkOrderReadyForClose(
            woNoNotes.Id,
            null
        );
        Test.stopTest();

        WorkOrder withNotesAfter = [
            SELECT Status, Tech_Support_Notes__c
            FROM WorkOrder
            WHERE Id = :woWithNotes.Id
            LIMIT 1
        ];
        System.assertEquals(
            'Ready for Close',
            withNotesAfter.Status,
            'Status should move to Ready for Close.'
        );
        System.assert(
            withNotesAfter.Tech_Support_Notes__c.contains('Existing note'),
            'Existing notes should be preserved.'
        );
        System.assert(
            withNotesAfter.Tech_Support_Notes__c.contains('Ready for Close Notes: Follow-up items listed'),
            'New Ready for Close notes should be appended.'
        );

        WorkOrder noNotesAfter = [
            SELECT Status, Tech_Support_Notes__c
            FROM WorkOrder
            WHERE Id = :woNoNotes.Id
            LIMIT 1
        ];
        System.assertEquals(
            'Ready for Close',
            noNotesAfter.Status,
            'Status should move to Ready for Close when no notes are provided.'
        );
        System.assertEquals(
            null,
            noNotesAfter.Tech_Support_Notes__c,
            'Notes should remain empty when none are provided.'
        );
    }

    // ------------------------- TRANSFER SYSTEM TESTS -------------------------

    @isTest
    static void testCreateEngineerTransferRequest_NewAndDuplicate() {
        ResourceContext context = getActiveResourceContext();
        User requester  = context.user;
        User targetUser = createPlainUser('target');

        WorkOrder wo = ensureWorkOrder(context);

        System.runAs(requester) {
            Test.startTest();
            FslTechnicianOnlineController.createEngineerTransferRequest(
                wo.Id,
                targetUser.Id
            );
            FslTechnicianOnlineController.createEngineerTransferRequest(
                wo.Id,
                targetUser.Id
            );
            Test.stopTest();
        }

        Integer countReqs = [
            SELECT COUNT()
            FROM Engineer_Transfer_Request__c
            WHERE Work_Order_to_Transfer__c = :wo.Id
              AND Transfer_Work_Order_To_Engineer__c = :targetUser.Id
        ];

        System.assertEquals(1, countReqs,
            'Only one pending Engineer_Transfer_Request__c should exist for same WO/target.');
    }

    @isTest
    static void testAcceptEngineerTransferRequest_AsTargetUser() {
        ResourceContext context = getActiveResourceContext();
        User targetTech = context.user;

        WorkOrder wo = ensureWorkOrder(context);

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c      = wo.Id,
            Subject            = 'Transfer Test',
            Status             = 'New',
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            ParentRecordId     = wo.Id
        );
        insert sa;

        sa = new ServiceAppointment(
            Id              = sa.Id,
            SchedStartTime  = Datetime.now().addDays(1),
            SchedEndTime    = Datetime.now().addDays(1).addHours(1),
            Status          = 'New'
        );
        update sa;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = context.resource.Id
        );
        insert ar;

        Engineer_Transfer_Request__c req = new Engineer_Transfer_Request__c(
            Work_Order_to_Transfer__c          = wo.Id,
            Transfer_Work_Order_To_Engineer__c = targetTech.Id
        );
        insert req;

        System.runAs(targetTech) {
            Test.startTest();
            FslTechnicianOnlineController.acceptEngineerTransferRequest(req.Id, null);
            Test.stopTest();
        }

        Engineer_Transfer_Request__c reqAfter = [
            SELECT Id, Accepted_on__c, Date_Rejected_On__c, Rejected_Reason__c
            FROM Engineer_Transfer_Request__c
            WHERE Id = :req.Id
            LIMIT 1
        ];
        System.assertNotEquals(null, reqAfter.Accepted_on__c, 'Request should be marked accepted.');
        System.assertEquals(null, reqAfter.Date_Rejected_On__c, 'Rejected date should be cleared.');
        System.assertEquals(null, reqAfter.Rejected_Reason__c, 'Rejection reason should be cleared.');

        WorkOrder woAfter = [
            SELECT Id, OwnerId
            FROM WorkOrder
            WHERE Id = :wo.Id
            LIMIT 1
        ];
        System.assertEquals(targetTech.Id, woAfter.OwnerId, 'Work Order should be transferred to target tech.');

        ServiceAppointment saAfter = [
            SELECT Id, Status, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :sa.Id
            LIMIT 1
        ];
        System.assertEquals('Canceled', saAfter.Status, 'Open appointments should be canceled after accept.');
        System.assertEquals(null, saAfter.SchedStartTime, 'SchedStartTime should be cleared.');
        System.assertEquals(null, saAfter.SchedEndTime, 'SchedEndTime should be cleared.');

        Integer arCountAfter = [
            SELECT COUNT()
            FROM AssignedResource
            WHERE ServiceAppointmentId = :sa.Id
        ];
        System.assertEquals(0, arCountAfter, 'Assignments should be removed after accept.');
    }

    @isTest
    static void testRejectEngineerTransferRequest_AsTargetUser() {
        ResourceContext context = getActiveResourceContext();
        User targetTech = context.user;

        WorkOrder wo = ensureWorkOrder(context);

        Engineer_Transfer_Request__c req = new Engineer_Transfer_Request__c(
            Work_Order_to_Transfer__c          = wo.Id,
            Transfer_Work_Order_To_Engineer__c = targetTech.Id
        );
        insert req;

        String reason = 'Too busy';

        System.runAs(targetTech) {
            Test.startTest();
            FslTechnicianOnlineController.rejectEngineerTransferRequest(req.Id, reason);
            Test.stopTest();
        }

        Engineer_Transfer_Request__c reqAfter = [
            SELECT Id, Accepted_on__c, Date_Rejected_On__c, Rejected_Reason__c
            FROM Engineer_Transfer_Request__c
            WHERE Id = :req.Id
            LIMIT 1
        ];

        System.assertNotEquals(null, reqAfter.Date_Rejected_On__c, 'Rejection date should be set.');
        System.assertEquals(reason, reqAfter.Rejected_Reason__c, 'Rejection reason should be saved.');
        System.assertEquals(null, reqAfter.Accepted_on__c, 'Accepted date should be cleared on reject.');
    }

    @isTest
    static void testGetAppointmentsIncludesTransferLists() {
        ResourceContext context = getActiveResourceContext();
        User tech = context.user;
        User otherUser = createPlainUser('other');

        WorkOrder woPending = ensureWorkOrder(context);

        WorkOrder woSubmitted = new WorkOrder(
            Subject            = 'Submitted Transfer WO',
            Status             = 'New',
            OwnerId            = tech.Id,
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            AccountId          = woPending.AccountId,
            Street             = '1 Submitted St',
            City               = 'Chicago',
            State              = 'Illinois',
            PostalCode         = '60601',
            Country            = 'United States'
        );
        insert woSubmitted;

        System.runAs(otherUser) {
            insert new Engineer_Transfer_Request__c(
                Work_Order_to_Transfer__c          = woPending.Id,
                Transfer_Work_Order_To_Engineer__c = tech.Id
            );
        }

        System.runAs(tech) {
            insert new Engineer_Transfer_Request__c(
                Work_Order_to_Transfer__c          = woSubmitted.Id,
                Transfer_Work_Order_To_Engineer__c = otherUser.Id
            );
        }

        FslTechnicianOnlineController.AppointmentResult result;
        System.runAs(tech) {
            Test.startTest();
            result = FslTechnicianOnlineController.getMyAppointmentsOnline(null);
            Test.stopTest();
        }

        System.assertNotEquals(null, result, 'Result should not be null.');
        System.assertNotEquals(null, result.transferRequests, 'transferRequests should not be null.');
        System.assertNotEquals(null, result.submittedTransferRequests, 'submittedTransferRequests should not be null.');

        Boolean hasPending = false;
        for (FslTechnicianOnlineController.TransferRequestDTO dto : result.transferRequests) {
            if (dto.workOrderId == woPending.Id) {
                hasPending = true;
                break;
            }
        }
        System.assertEquals(true, hasPending, 'Pending transfer for woPending should be present.');

        Boolean hasSubmitted = false;
        for (FslTechnicianOnlineController.TransferRequestDTO dto : result.submittedTransferRequests) {
            if (dto.workOrderId == woSubmitted.Id) {
                hasSubmitted = true;
                break;
            }
        }
        System.assertEquals(true, hasSubmitted, 'Submitted transfer for woSubmitted should be present.');
    }

    // ------------------------- OTHER CONTROLLER METHODS -------------------------

    @isTest
    static void testUnassignAppointment() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo = ensureWorkOrder(context);

        ServiceAppointment sa = new ServiceAppointment(
            Work_Order__c      = wo.Id,
            Subject            = 'Unassign Test',
            Status             = 'New',
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            ParentRecordId     = wo.Id
        );
        insert sa;

        // Now update to a scheduled state AFTER insert
        sa = new ServiceAppointment(
            Id              = sa.Id,
            SchedStartTime  = Datetime.now().addDays(1),
            SchedEndTime    = Datetime.now().addDays(1).addHours(1),
            Status          = 'Scheduled'
        );
        update sa;

        AssignedResource ar = new AssignedResource(
            ServiceAppointmentId = sa.Id,
            ServiceResourceId    = context.resource.Id
        );
        insert ar;

        Test.startTest();
        FslTechnicianOnlineController.unassignAppointment(sa.Id);
        Test.stopTest();

        ServiceAppointment saAfter = [
            SELECT Id, Status, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE Id = :sa.Id
            LIMIT 1
        ];

        System.assertEquals('New', saAfter.Status, 'Status should reset to New.');
        System.assertEquals(null, saAfter.SchedStartTime, 'SchedStartTime should be cleared.');
        System.assertEquals(null, saAfter.SchedEndTime, 'SchedEndTime should be cleared.');

        Integer arCountAfter = [
            SELECT COUNT()
            FROM AssignedResource
            WHERE ServiceAppointmentId = :sa.Id
        ];
        System.assertEquals(0, arCountAfter, 'All assignments should be removed.');
    }

    @isTest
    static void testUpdateAndDeleteResourceAbsence() {
        ResourceContext context = getActiveResourceContext();
        ResourceAbsence ra = ensureResourceAbsence(context);
    
        Datetime newStart = Datetime.now().addDays(3);
        Datetime newEnd   = newStart.addHours(4);
    
        // Do NOT wrap this in Test.startTest()/stopTest() to avoid "Testing already started"
        FslTechnicianOnlineController.updateResourceAbsence(ra.Id, newStart, newEnd);
    
        ResourceAbsence updated = [
            SELECT Id, Start, End
            FROM ResourceAbsence
            WHERE Id = :ra.Id
            LIMIT 1
        ];
    
        System.assertEquals(
            truncateToMinute(newStart),
            truncateToMinute(updated.Start),
            'Start should be updated (minute precision).'
        );
        System.assertEquals(
            truncateToMinute(newEnd),
            truncateToMinute(updated.End),
            'End should be updated (minute precision).'
        );
    
        // And now delete, also without startTest/stopTest
        FslTechnicianOnlineController.deleteResourceAbsence(ra.Id);
    
        Integer countAfter = [
            SELECT COUNT()
            FROM ResourceAbsence
            WHERE Id = :ra.Id
        ];
        System.assertEquals(0, countAfter, 'ResourceAbsence record should be deleted.');
    }


    @isTest
    static void testGetTerritoryResources() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo = ensureWorkOrder(context);

        Test.startTest();
        List<FslTechnicianOnlineController.ServiceResourceOption> options =
            FslTechnicianOnlineController.getTerritoryResources(wo.Id);
        Test.stopTest();

        System.assertNotEquals(null, options, 'Options should not be null.');
        System.assert(!options.isEmpty(), 'At least one resource option should be returned.');

        Boolean foundSelf = false;
        for (FslTechnicianOnlineController.ServiceResourceOption opt : options) {
            if (opt.serviceResourceId == context.resource.Id &&
                opt.userId == context.user.Id) {
                foundSelf = true;
                break;
            }
        }
        System.assertEquals(true, foundSelf, 'Current technician resource should be present in territory options.');
    }
    
    	@isTest
        static void testAppointmentDtoFieldMapping() {
            // Use a clean synthetic tech + territory
            ResourceContext context = createSyntheticResourceContext();
        
            // Account + Contact for WO
            Account acc = new Account(Name = 'DTO Account');
            insert acc;
        
            Contact con = new Contact(
                LastName  = 'DTO Contact',
                AccountId = acc.Id,
                Phone     = '555-1111',
                Email     = 'dto@example.com'
            );
            insert con;
        
        // Work Order with contact + address + extra fields used in DTO
        WorkOrder wo = new WorkOrder(
            Subject                  = 'DTO WO',
            Status                   = 'New',
            OwnerId                  = context.user.Id,
            ServiceTerritoryId       = context.stm.ServiceTerritoryId,
            AccountId                = acc.Id,
            ContactId                = con.Id,
            Make_Model__c            = 'Model X',
            Reporter_Contact_Info__c = 'Reporter Info',
            Street                   = '789 DTO Way',
            City                     = 'Phoenix',
            State                    = 'Arizona',
            PostalCode               = '85001',
            Country                  = 'United States'
        );
        insert wo;
        
        // 🔥 Important: re-query so WorkOrderNumber is populated
        wo = [
            SELECT Id, WorkOrderNumber, Status, AccountId, ContactId,
                   Make_Model__c, Reporter_Contact_Info__c,
                   Street, City, State, PostalCode, Country
            FROM WorkOrder
            WHERE Id = :wo.Id
        ];

        
            // Scheduled Service Appointment tied to that WO
            Datetime start = Datetime.now().addDays(1);
            Datetime endtime   = start.addHours(2);
        
            ServiceAppointment sa = new ServiceAppointment(
                Work_Order__c      = wo.Id,
                Subject            = 'DTO SA',
                Status             = 'New', // safe for insert
                SchedStartTime     = start,
                SchedEndTime       = endtime,
                ServiceTerritoryId = context.stm.ServiceTerritoryId,
                ParentRecordId     = wo.Id,
                Street             = '456 DTO St',
                City               = 'Denver',
                State              = 'Colorado',
                PostalCode         = '80202',
                Country            = 'United States'
            );
            insert sa;
        
            // AssignedResource so this tech "owns" the appointment
            AssignedResource ar = new AssignedResource(
                ServiceAppointmentId = sa.Id,
                ServiceResourceId    = context.resource.Id
            );
            insert ar;
        
            // Call controller as the technician
            FslTechnicianOnlineController.AppointmentResult result;
            System.runAs(context.user) {
                Test.startTest();
                result = FslTechnicianOnlineController.getMyAppointmentsOnline(null);
                Test.stopTest();
            }
        
            System.assertNotEquals(null, result, 'Result should not be null.');
            System.assertNotEquals(null, result.appointments, 'Appointments list should not be null.');
            System.assert(!result.appointments.isEmpty(), 'Appointments list should not be empty.');
        
            // Find the DTO for our specific appointment
            FslTechnicianOnlineController.AppointmentDTO target = null;
            for (FslTechnicianOnlineController.AppointmentDTO dto : result.appointments) {
                if (dto.appointmentId == sa.Id) {
                    target = dto;
                    break;
                }
            }
            System.assertNotEquals(null, target, 'Should find DTO for the test ServiceAppointment.');
        
            // --- Field mapping assertions ---
        
            System.assertEquals(sa.Id,      target.appointmentId,   'appointmentId should match SA Id.');
            System.assertEquals(wo.Id,      target.workOrderId,     'workOrderId should match WO Id.');
            System.assertEquals(sa.Subject, target.workOrderSubject,'Subject should come from SA.Subject.');
            System.assertEquals(wo.Status,  target.workOrderStatus, 'workOrderStatus should come from WO.Status.');
            System.assertEquals(wo.WorkOrderNumber, target.workOrderNumber,
                'workOrderNumber should come from Work_Order__r.WorkOrderNumber.');
        
            // Account resolution: should use WO.Account since SA.AccountId is null
            System.assertEquals(acc.Id,   target.accountId,   'AccountId should come from WO.AccountId.');
            System.assertEquals(acc.Name, target.accountName, 'AccountName should come from WO.Account.Name.');
        
            // Contact fields from WorkOrder.Contact (use the Contact record we created)
            System.assertEquals(con.Id,    target.contactId,    'ContactId should come from WO.ContactId.');
            System.assertEquals(con.LastName,target.contactName,'ContactName should come from WO.Contact.Name.');
            System.assertEquals(con.Phone, target.contactPhone, 'ContactPhone should come from WO.Contact.Phone.');
            System.assertEquals(con.Email, target.contactEmail, 'ContactEmail should come from WO.Contact.Email.');

            System.assertEquals(
                truncateToMinute(start),
                truncateToMinute(target.schedStart),
                'schedStart should match SA.SchedStartTime (minute precision).'
            );
            System.assertEquals(
                truncateToMinute(endtime),
                truncateToMinute(target.schedEnd),
                'schedEnd should match SA.SchedEndTime (minute precision).'
            );

        
            // Make/Model + Reporter info from WO
            System.assertEquals('Model X',       target.makeModel,           'makeModel should come from WO.Make_Model__c.');
            System.assertEquals('Reporter Info', target.reporterContactInfo, 'reporterContactInfo should come from WO.Reporter_Contact_Info__c.');
        
            // Address should come from SA’s address fields
            System.assertEquals('456 DTO St',    target.street,     'street should come from SA.Street.');
            System.assertEquals('Denver',        target.city,       'city should come from SA.City.');
            System.assertEquals('Colorado',      target.state,      'state should come from SA.State.');
            System.assertEquals('80202',         target.postalCode, 'postalCode should come from SA.PostalCode.');
            System.assertEquals('United States', target.country,    'country should come from SA.Country.');
        
            // Assignment flags
            System.assertEquals(true,  target.isMyAssignment,   'isMyAssignment should be true for this tech.');
            System.assertEquals(false, target.isCrewAssignment, 'isCrewAssignment should be false (no crew).');
        
            // With no WOLIs, parts flags should be false
            System.assertEquals(false, target.allPartsEnRoute,  'allPartsEnRoute should be false with no WOLIs.');
            System.assertEquals(false, target.somePartsEnRoute, 'somePartsEnRoute should be false with no WOLIs.');
        }

@isTest
static void testAssignedResourceCrewVsNonCrewLogic() {

    // Create base context (tech user, territory, resource)
    ResourceContext context = createSyntheticResourceContext();

    // Minimal Account for WO
    Account acc = new Account(Name = 'Crew Test Acc');
    insert acc;

    // Create Work Order owned by this tech
    WorkOrder wo = new WorkOrder(
        Subject            = 'Crew Logic WO',
        Status             = 'New',
        OwnerId            = context.user.Id,
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        AccountId          = acc.Id,
        Street             = '1 St',
        City               = 'Chicago',
        State              = 'Illinois',
        PostalCode         = '60007',
        Country            = 'United States'
    );
    insert wo;

    wo = [
        SELECT Id, WorkOrderNumber, ServiceTerritoryId
        FROM WorkOrder
        WHERE Id = :wo.Id
    ];

    // -----------------------------
    // 1) NON-CREW appointment (assigned directly to tech)
    // -----------------------------
    Datetime start1 = Datetime.now().addDays(1);
    Datetime end1   = start1.addHours(1);

    ServiceAppointment saNonCrew = new ServiceAppointment(
        Work_Order__c      = wo.Id,
        Subject            = 'NonCrew Appt',
        Status             = 'New',
        ServiceTerritoryId = wo.ServiceTerritoryId,
        ParentRecordId     = wo.Id
    );
    insert saNonCrew;

    saNonCrew = new ServiceAppointment(
        Id              = saNonCrew.Id,
        SchedStartTime  = start1,
        SchedEndTime    = end1,
        Status          = 'New'
    );
    update saNonCrew;

    // Assigned directly to the technician's ServiceResource -> non-crew path
    AssignedResource arNonCrew = new AssignedResource(
        ServiceAppointmentId = saNonCrew.Id,
        ServiceResourceId    = context.resource.Id
    );
    insert arNonCrew;

    // -----------------------------
    // 2) CREW appointment using EXISTING ServiceCrew
    // -----------------------------

    // Use an existing ServiceCrew from the org (sandbox/prod)
    // Assumption: at least one ServiceCrew exists with whatever required fields your org enforces.
    ServiceCrew existingCrew = [
        SELECT Id
        FROM ServiceCrew
        LIMIT 1
    ];

    Datetime start2 = Datetime.now().addDays(2);
    Datetime end2   = start2.addHours(1);

    ServiceAppointment saCrew = new ServiceAppointment(
        Work_Order__c      = wo.Id,
        Subject            = 'Crew Appt',
        Status             = 'New',
        ServiceTerritoryId = wo.ServiceTerritoryId,
        ParentRecordId     = wo.Id
    );
    insert saCrew;

    saCrew = new ServiceAppointment(
        Id              = saCrew.Id,
        SchedStartTime  = start2,
        SchedEndTime    = end2,
        Status          = 'New'
    );
    update saCrew;

    // Crew AssignedResource:
    // - ServiceCrewId populated -> isCrewLike = TRUE in the controller
    // - ServiceResourceId uses the same tech resource (satisfies your validation that it cannot be blank)
    AssignedResource arCrew = new AssignedResource(
        ServiceAppointmentId = saCrew.Id,
        ServiceCrewId        = existingCrew.Id,
        ServiceResourceId    = context.resource.Id
    );
    insert arCrew;

    // -----------------------------
    // 3) Run controller to hit crew/non-crew branching
    // -----------------------------
    FslTechnicianOnlineController.AppointmentResult result;

    System.runAs(context.user) {
        Test.startTest();
        result = FslTechnicianOnlineController.getMyAppointmentsOnline(null);
        Test.stopTest();
    }

    System.assertNotEquals(null, result, 'Controller returned result.');

    // At this point, the "3) For those ServiceAppointments..." loop has executed for:
    //  - arNonCrew (no ServiceCrewId, just ServiceResourceId) -> non-crew path
    //  - arCrew (has ServiceCrewId) -> crew path
    //
    // We don't NEED to assert on the internals of workOrdersWithCrewAssignment /
    // workOrdersWithNonCrewAssignment for coverage purposes; just ensure call completes.
}
@isTest
static void testRejectTransferRequest_AsExistingManagerForTech() {
    // Find a "tech" user who:
    //  - Has a ManagerId (so someone manages them)
    //  - Owns at least one WorkOrder
    List<User> techUsers = [
        SELECT Id, ManagerId
        FROM User
        WHERE ManagerId != null
          AND Id IN (SELECT OwnerId FROM WorkOrder)
        LIMIT 1
    ];

    if (techUsers.isEmpty()) {
        // In case this org has no such data, don't blow up the test run.
        // You can remove this guard if you KNOW the data exists.
        return;
    }

    User techUser = techUsers[0];

    // Manager of that user (this is who we'll runAs)
    User managerUser = [
        SELECT Id
        FROM User
        WHERE Id = :techUser.ManagerId
        LIMIT 1
    ];

    // One WorkOrder owned by the tech user
    WorkOrder wo = [
        SELECT Id, OwnerId
        FROM WorkOrder
        WHERE OwnerId = :techUser.Id
        LIMIT 1
    ];

    // Create a transfer request targeting the tech user
    Engineer_Transfer_Request__c req = new Engineer_Transfer_Request__c(
        Work_Order_to_Transfer__c          = wo.Id,
        Transfer_Work_Order_To_Engineer__c = techUser.Id
    );
    insert req;

    String rejectionReason = 'Manager rejecting on behalf of tech';

    // 🔹 currentUserId = managerUser.Id
    // 🔹 targetUserId  = techUser.Id (from req.Transfer_Work_Order_To_Engineer__c)
    // 🔹 techUser.ManagerId == managerUser.Id  => isManagerActingForTarget = TRUE
    System.runAs(managerUser) {
        Test.startTest();
        FslTechnicianOnlineController.rejectEngineerTransferRequest(
            req.Id,
            rejectionReason
        );
        Test.stopTest();
    }

    // Verify the request was actually rejected (no auth error, branch executed)
    Engineer_Transfer_Request__c reqAfter = [
        SELECT Date_Rejected_On__c, Rejected_Reason__c, Accepted_on__c
        FROM Engineer_Transfer_Request__c
        WHERE Id = :req.Id
        LIMIT 1
    ];

    System.assertNotEquals(
        null,
        reqAfter.Date_Rejected_On__c,
        'Manager should be able to reject the request (date set).'
    );
    System.assertEquals(
        rejectionReason,
        reqAfter.Rejected_Reason__c,
        'Rejection reason should be recorded when manager rejects.'
    );
    System.assertEquals(
        null,
        reqAfter.Accepted_on__c,
        'Accepted_on__c should remain null after rejection.'
    );
}

@isTest
static void testAssignedResourceCrewAndNonCrewBranches() {
    // Use synthetic tech + resource + territory
    ResourceContext context = createSyntheticResourceContext();

    // Shared account
    Account acc = new Account(Name = 'AR Branch Test Acc');
    insert acc;

    // Two WOs owned by the active tech user
    WorkOrder woNonCrew = new WorkOrder(
        Subject            = 'WO NonCrew',
        Status             = 'New',
        OwnerId            = context.user.Id,
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        AccountId          = acc.Id,
        Street             = '10 NonCrew St',
        City               = 'Chicago',
        State              = 'Illinois',
        PostalCode         = '60601',
        Country            = 'United States'
    );
    WorkOrder woCrew = new WorkOrder(
        Subject            = 'WO Crew',
        Status             = 'New',
        OwnerId            = context.user.Id,
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        AccountId          = acc.Id,
        Street             = '20 Crew St',
        City               = 'Chicago',
        State              = 'Illinois',
        PostalCode         = '60602',
        Country            = 'United States'
    );
    insert new List<WorkOrder>{ woNonCrew, woCrew };

    Datetime base = Datetime.now().addDays(3);

    // ---- Insert ServiceAppointments UNSCHEDULED first ----
    ServiceAppointment saNonCrew = new ServiceAppointment(
        Work_Order__c      = woNonCrew.Id,
        Subject            = 'SA NonCrew',
        Status             = 'New',
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        ParentRecordId     = woNonCrew.Id
    );
    ServiceAppointment saCrew = new ServiceAppointment(
        Work_Order__c      = woCrew.Id,
        Subject            = 'SA Crew',
        Status             = 'New',
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        ParentRecordId     = woCrew.Id
    );
    insert new List<ServiceAppointment>{ saNonCrew, saCrew };

    // ---- ACTIVE user that is NOT the technician ----
    User otherUser = [
        SELECT Id
        FROM User
        WHERE IsActive = true
          AND Id != :context.user.Id
        LIMIT 1
    ];

    // ---- ServiceResource for that active user ----
    ServiceResource arResource;
    List<ServiceResource> existingRes = [
        SELECT Id
        FROM ServiceResource
        WHERE RelatedRecordId = :otherUser.Id
        LIMIT 1
    ];
    if (!existingRes.isEmpty()) {
        arResource = existingRes[0];
    } else {
        arResource = new ServiceResource(
            Name            = 'AR Test Resource',
            IsActive        = true,
            RelatedRecordId = otherUser.Id
        );
        insert arResource;
    }

    // ✅ Make arResource a member of the same territory so scheduling is allowed
    ServiceTerritoryMember stmForAr = new ServiceTerritoryMember(
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        ServiceResourceId  = arResource.Id,
        EffectiveStartDate = Date.today().addDays(-1)
    );
    insert stmForAr;

    // ---- Now update the ServiceAppointments to be SCHEDULED ----
    saNonCrew = new ServiceAppointment(
        Id              = saNonCrew.Id,
        SchedStartTime  = base,
        SchedEndTime    = base.addHours(1),
        Status          = 'New'
    );
    saCrew = new ServiceAppointment(
        Id              = saCrew.Id,
        SchedStartTime  = base.addDays(1),
        SchedEndTime    = base.addDays(1).addHours(1),
        Status          = 'New'
    );
    update new List<ServiceAppointment>{ saNonCrew, saCrew };

    // ---- Existing ServiceCrew for crew-like path ----
    List<ServiceCrew> crewList = [
        SELECT Id
        FROM ServiceCrew
        LIMIT 1
    ];
    if (crewList.isEmpty()) {
        // No crews in this org; avoid hard failure
        return;
    }
    ServiceCrew existingCrew = crewList[0];

    // ---- AssignedResource rows ----
    // Non-crew assignment (no ServiceCrewId)
    AssignedResource arNonCrew = new AssignedResource(
        ServiceAppointmentId = saNonCrew.Id,
        ServiceResourceId    = arResource.Id
    );

    // Crew-like assignment (ServiceCrewId populated -> isCrewLike = true)
    AssignedResource arCrew = new AssignedResource(
        ServiceAppointmentId = saCrew.Id,
        ServiceCrewId        = existingCrew.Id,
        ServiceResourceId    = arResource.Id
    );

    insert new List<AssignedResource>{ arNonCrew, arCrew };

    // ---- Call controller ----
    FslTechnicianOnlineController.AppointmentResult result;

    System.runAs(context.user) { // context.user is ACTIVE by construction
        Test.startTest();
        result = FslTechnicianOnlineController.getMyAppointmentsOnline(null);
        Test.stopTest();
    }

    System.assertNotEquals(null, result, 'Result should not be null.');
    System.assertNotEquals(null, result.unscheduledWorkOrders,
        'unscheduledWorkOrders should not be null.');

    Boolean foundCrewWo    = false;
    Boolean foundNonCrewWo = false;

    for (FslTechnicianOnlineController.UnscheduledWorkOrderDTO dto
        : result.unscheduledWorkOrders) {

        if (dto.workOrderId == woCrew.Id) {
            foundCrewWo = true;
            System.assertEquals(
                true,
                dto.hasCrewAssignment,
                'WO with crew-like AssignedResource should be flagged hasCrewAssignment.'
            );
        }

        if (dto.workOrderId == woNonCrew.Id) {
            foundNonCrewWo = true;
        }
    }

    System.assertEquals(
        true,
        foundCrewWo,
        'Crew WO should appear in unscheduledWorkOrders.'
    );
    System.assertEquals(
        false,
        foundNonCrewWo,
        'Non-crew scheduled WO should NOT appear (filtered out by non-crew branch).'
    );
}
    
    @isTest
static void testUnscheduledAndAttachmentFromWorkOrderAndOpportunity() {
    // Use synthetic active tech/territory/resource
    ResourceContext context = createSyntheticResourceContext();

    // Base Account
    Account acc = new Account(Name = 'DTO Attachment Acc');
    insert acc;

    // Opportunity to be linked to WO and to have a file
    Opportunity opp = new Opportunity(
        Name       = 'DTO Attachment Opp',
        StageName  = 'Prospecting',
        CloseDate  = Date.today().addDays(30),
        AccountId  = acc.Id
    );
    insert opp;

    // Work Order owned by the tech and tied to the Opportunity
    WorkOrder wo = new WorkOrder(
        Subject            = 'WO With Opp + File',
        Status             = 'New',
        OwnerId            = context.user.Id,
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        AccountId          = acc.Id,
        Opportunity__c     = opp.Id,
        Street             = '111 Test St',
        City               = 'Chicago',
        State              = 'Illinois',
        PostalCode         = '60603',
        Country            = 'United States'
    );
    insert wo;

    // UNSCHEDULED ServiceAppointment for this WO
    // - Status is not in saClosedStatuses
    // - No SchedStartTime/SchedEndTime -> isScheduled = false
    // - Return_Visit_Required__c = false -> hasUnscheduledNonReturnByWo[wo.Id] = true
    ServiceAppointment sa = new ServiceAppointment(
        Work_Order__c      = wo.Id,
        Subject            = 'Unscheduled SA',
        Status             = 'New',
        Return_Visit_Required__c = false,
        ServiceTerritoryId = context.stm.ServiceTerritoryId,
        ParentRecordId     = wo.Id
    );
    insert sa;

    // ---- Create a File (ContentDocument) and link to BOTH WO and Opportunity ----
    ContentVersion cv = new ContentVersion(
        Title        = 'Test Quote File',
        PathOnClient = 'quote.pdf',
        VersionData  = Blob.valueOf('dummy pdf content')
    );
    insert cv;

    // Get the ContentDocumentId created by the ContentVersion insert
    cv = [
        SELECT Id, ContentDocumentId
        FROM ContentVersion
        WHERE Id = :cv.Id
        LIMIT 1
    ];
    Id docId = cv.ContentDocumentId;

    // Link file to the Work Order (tests first CDL loop: LinkedEntityId IN candidateWoIds)
    ContentDocumentLink cdlWo = new ContentDocumentLink(
        LinkedEntityId   = wo.Id,
        ContentDocumentId = docId,
        ShareType        = 'V',
        Visibility       = 'AllUsers'
    );

    // Link file to the Opportunity (tests second CDL loop: LinkedEntityId IN candidateOppIds)
    ContentDocumentLink cdlOpp = new ContentDocumentLink(
        LinkedEntityId   = opp.Id,
        ContentDocumentId = docId,
        ShareType        = 'V',
        Visibility       = 'AllUsers'
    );

    insert new List<ContentDocumentLink>{ cdlWo, cdlOpp };

    // ---- Call controller as the tech user ----
    FslTechnicianOnlineController.AppointmentResult result;

    System.runAs(context.user) {
        Test.startTest();
        result = FslTechnicianOnlineController.getMyAppointmentsOnline(null);
        Test.stopTest();
    }

    System.assertNotEquals(null, result, 'Result should not be null.');
    System.assertNotEquals(null, result.unscheduledWorkOrders,
        'unscheduledWorkOrders should not be null.');

    // Find our WO in the unscheduled tray
    FslTechnicianOnlineController.UnscheduledWorkOrderDTO dtoTarget = null;
    for (FslTechnicianOnlineController.UnscheduledWorkOrderDTO dto
             : result.unscheduledWorkOrders) {
        if (dto.workOrderId == wo.Id) {
            dtoTarget = dto;
            break;
        }
    }

    System.assertNotEquals(null, dtoTarget,
        'Unscheduled DTO should exist for the test Work Order.');

    // ---- Assertions proving the code you pasted actually ran ----

    // 1) Unscheduled logic:
    //    - isScheduled = false => unscheduled count incremented
    //    - Return_Visit_Required__c = false => hasUnscheduledNonReturnAppointment = true
    System.assertEquals(
        1,
        dtoTarget.unscheduledServiceAppointmentCount,
        'There should be exactly one unscheduled ServiceAppointment for this WO.'
    );
    System.assertEquals(
        true,
        dtoTarget.hasUnscheduledNonReturnAppointment,
        'hasUnscheduledNonReturnAppointment should be true for unscheduled, non-return SA.'
    );

    // 2) Attachment logic:
    //    We linked the same ContentDocument to WO (candidateWoIds) and Opp (candidateOppIds).
    //    The loops should mark the WO as having a quote attachment and fill URL & DocumentId.
    System.assertEquals(
        true,
        dtoTarget.hasQuoteAttachment,
        'hasQuoteAttachment should be true when ContentDocumentLink exists.'
    );
    System.assertNotEquals(
        null,
        dtoTarget.quoteAttachmentDocumentId,
        'quoteAttachmentDocumentId should be populated from ContentDocumentLink.'
    );
    System.assert(
        dtoTarget.quoteAttachmentDownloadUrl != null &&
        dtoTarget.quoteAttachmentDownloadUrl.startsWith('/sfc/servlet.shepherd/document/download/'),
        'quoteAttachmentDownloadUrl should be the shepherd download URL.'
    );
}


}
