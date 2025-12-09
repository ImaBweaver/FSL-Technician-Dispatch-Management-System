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
     * Build a synthetic technician, resource, and territory when live data is not found.
     */
    private static ResourceContext createSyntheticResourceContext() {
        Profile standardProfile = [SELECT Id FROM Profile LIMIT 1];

        String uniqueSuffix = String.valueOf(Datetime.now().getTime());
        User user = new User(
            ProfileId = standardProfile.Id,
            LastName = 'Tech',
            Alias = 'tech',
            Email = 'tech' + uniqueSuffix + '@example.com',
            Username = 'tech' + uniqueSuffix + '@example.com',
            TimeZoneSidKey = 'America/Los_Angeles',
            LocaleSidKey = 'en_US',
            EmailEncodingKey = 'UTF-8',
            LanguageLocaleKey = 'en_US'
        );
        insert user;

        ServiceTerritory territory = new ServiceTerritory(
            Name = 'Test Territory ' + uniqueSuffix,
            IsActive = true
        );
        insert territory;

        ServiceResource resource = new ServiceResource(
            Name = 'Tech Resource ' + uniqueSuffix,
            IsActive = true,
            RelatedRecordId = user.Id
        );
        insert resource;

        ServiceTerritoryMember stm = new ServiceTerritoryMember(
            ServiceTerritoryId = territory.Id,
            ServiceResourceId = resource.Id,
            EffectiveStartDate = Date.today().addDays(-1)
        );
        insert stm;

        ResourceContext context = new ResourceContext();
        context.stm = stm;
        context.resource = resource;
        context.user = user;

        return context;
    }

    /**
     * Locate an active Service Resource with an active territory assignment and matching Work Orders.
     * This helps ensure downstream tests work with consistent real data.
     */
    private static ResourceContext getActiveResourceContext() {
        List<ServiceTerritoryMember> stms = [
        ServiceTerritoryMember stm = [
            SELECT Id, ServiceResourceId, ServiceTerritoryId, ServiceResource.RelatedRecordId
            FROM ServiceTerritoryMember
            WHERE ServiceResource.RelatedRecordId != null
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

        if (stms.isEmpty()) {
            return createSyntheticResourceContext();
        }

        ServiceTerritoryMember stm = stms[0];
        ServiceResource resource = [
            SELECT Id, RelatedRecordId
            FROM ServiceResource
            WHERE Id = :stm.ServiceResourceId
            LIMIT 1
        ];

        User user = [SELECT Id FROM User WHERE Id = :resource.RelatedRecordId LIMIT 1];

        ResourceContext context = new ResourceContext();
        context.stm = stm;
        context.resource = resource;
        context.user = user;

        return context;
    }

    private static WorkOrder ensureWorkOrder(ResourceContext context) {
        List<WorkOrder> workOrders = [
            SELECT Id, OwnerId, ServiceTerritoryId, Status, Subject, AccountId, Street, City, State, PostalCode, Country
            FROM WorkOrder
            WHERE OwnerId = :context.user.Id
              AND ServiceTerritoryId = :context.stm.ServiceTerritoryId
            LIMIT 1
        ];

        if (!workOrders.isEmpty()) {
            return workOrders[0];
        }

        Account acc = new Account(Name = 'Test Account');
        insert acc;

        String woStatus = firstPicklistValue(WorkOrder.Status);
        if (woStatus == null) {
            woStatus = 'New';
        }
        WorkOrder wo = new WorkOrder(
            Subject = 'Test Work Order',
            Status = woStatus,
            OwnerId = context.user.Id,
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            AccountId = acc.Id,
            Street = '123 Main St',
            City = 'San Francisco',
            State = 'CA',
            PostalCode = '94105',
            Country = 'USA'
        );
        insert wo;

        return wo;
    }

    private static ServiceAppointment ensureScheduledAppointment(ResourceContext context) {
        List<ServiceAppointment> existing = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE ServiceTerritoryId = :context.stm.ServiceTerritoryId
              AND SchedStartTime != null AND SchedEndTime != null
            LIMIT 1
            FOR UPDATE
        ];

        if (!existing.isEmpty()) {
            return existing[0];
        }

        WorkOrder wo = ensureWorkOrder(context);
        Datetime startTime = Datetime.now().addDays(1);
        Datetime endTime = startTime.addHours(1);

        String saStatus = firstPicklistValue(ServiceAppointment.Status);
        if (saStatus == null) {
            saStatus = 'New';
        }
        String durationType = firstPicklistValue(ServiceAppointment.DurationType);
        if (durationType == null) {
            durationType = 'Minutes';
        }

        ServiceAppointment sa = new ServiceAppointment(
            Subject = 'Test Appointment',
            Status = saStatus,
            SchedStartTime = startTime,
            SchedEndTime = endTime,
            ServiceTerritoryId = context.stm.ServiceTerritoryId,
            ParentRecordId = wo.Id,
            Duration = 60,
            DurationType = durationType,
            Work_Order__c = wo.Id
        );
        insert sa;

        return sa;
        User user = [SELECT Id FROM User WHERE Id = :resource.RelatedRecordId LIMIT 1];

        ResourceContext context = new ResourceContext();
        context.stm = stm;
        context.resource = resource;
        context.user = user;

        return context;
    }

    /**
     * Locate a Service Resource and User that already exist in the org so we can
     * exercise the queries without inserting synthetic data.
     */
    private static User getExistingTechnicianUser() {
        return getActiveResourceContext().user;
    }

    @isTest
    static void testGetAppointmentsForActiveUser() {
        User techUser = getExistingTechnicianUser();

        System.runAs(techUser) {
            Test.startTest();
            FslTechnicianOnlineController.AppointmentResult result =
                FslTechnicianOnlineController.getMyAppointmentsOnline(null);
            Test.stopTest();

        }
    }

    @isTest
    static void testGetAppointmentsForOtherUser() {
        User targetUser = [
            SELECT Id
            FROM User
            WHERE IsActive = true
              AND Id != :UserInfo.getUserId()
            LIMIT 1
        ];

        Test.startTest();
        FslTechnicianOnlineController.AppointmentResult result =
            FslTechnicianOnlineController.getMyAppointmentsOnline(targetUser.Id);
        Test.stopTest();

    }

    @isTest
    static void testRescheduleAndEndTimeUpdate() {
        ResourceContext context = getActiveResourceContext();

        ServiceAppointment sa = ensureScheduledAppointment(context);
        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE ServiceTerritoryId = :context.stm.ServiceTerritoryId
              AND SchedStartTime != null AND SchedEndTime != null
            LIMIT 1
            FOR UPDATE
        ];

        Datetime newStart = sa.SchedStartTime.addDays(1);
        Datetime earlierEnd = sa.SchedStartTime.addMinutes(5);

        Test.startTest();
        FslTechnicianOnlineController.rescheduleAppointment(sa.Id, newStart);
        FslTechnicianOnlineController.updateAppointmentEnd(sa.Id, earlierEnd);
        Test.stopTest();
    }

    @isTest
    static void testAssignCrewAppointment() {
        ResourceContext context = getActiveResourceContext();

        ServiceAppointment sa = ensureScheduledAppointment(context);


        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE ServiceTerritoryId = :context.stm.ServiceTerritoryId
              AND SchedStartTime != null AND SchedEndTime != null
            LIMIT 1
            FOR UPDATE
        ];

        Test.startTest();
        FslTechnicianOnlineController.assignCrewAppointment(sa.Id, context.resource.Id);
        Test.stopTest();
    }

    @isTest
    static void testCreateAppointmentAndOptions() {
        ResourceContext context = getActiveResourceContext();
        User techUser = context.user;

        WorkOrder wo = ensureWorkOrder(context);
        WorkOrder wo = [
            SELECT Id, OwnerId, ServiceTerritoryId, Status, Subject, AccountId, Street, City, State, PostalCode, Country
            FROM WorkOrder
            WHERE OwnerId = :techUser.Id
              AND ServiceTerritoryId = :context.stm.ServiceTerritoryId
            LIMIT 1
        ];

        Datetime startTime = Datetime.now().addDays(2);
        Datetime endTime = startTime.addHours(1);

        Test.startTest();
        Id newApptId = FslTechnicianOnlineController.createAppointmentForWorkOrder(
            wo.Id,
            startTime,
            endTime,
            techUser.Id
        );
        Test.stopTest();



       
    }

    @isTest
    static void testRequestRescheduleAndMarkQuoteSent() {
        ResourceContext context = getActiveResourceContext();
        WorkOrder wo = ensureWorkOrder(context);
        ServiceTerritoryMember stm = [
            SELECT Id, ServiceResourceId, ServiceTerritoryId, ServiceResource.RelatedRecordId
            FROM ServiceTerritoryMember
            WHERE ServiceResource.RelatedRecordId != null
              AND EffectiveStartDate != null
              AND EffectiveStartDate <= :Date.today()
              AND (EffectiveEndDate = null OR EffectiveEndDate >= :Date.today())
              AND ServiceTerritoryId IN (
                  SELECT ServiceTerritoryId
                  FROM WorkOrder
                  WHERE ServiceTerritoryId != null
              )
            LIMIT 1
        ];

        ServiceResource resource = [
            SELECT Id, RelatedRecordId
            FROM ServiceResource
            WHERE Id = :stm.ServiceResourceId
            LIMIT 1
        ];

        WorkOrder wo = [
            SELECT Id, OwnerId
            FROM WorkOrder
            WHERE OwnerId = :resource.RelatedRecordId
              AND ServiceTerritoryId = :stm.ServiceTerritoryId
            LIMIT 1
        ];

        Test.startTest();
        FslTechnicianOnlineController.requestRescheduleToResource(wo.Id, context.resource.Id);
        FslTechnicianOnlineController.markWorkOrderQuoteSent(wo.Id);
        Test.stopTest();
    }
