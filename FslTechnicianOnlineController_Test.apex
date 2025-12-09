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
     * Locate an active Service Resource with an active territory assignment and matching Work Orders.
     * This helps ensure downstream tests work with consistent real data.
     */
    private static ResourceContext getActiveResourceContext() {
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

            System.assertNotEquals(null, result, 'Result should be returned for active user.');
            System.assertNotEquals(null, result.debug, 'Debug info should be populated.');
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

        System.assertNotEquals(null, result, 'Result should be returned when requesting another user.');
        System.assertEquals(targetUser.Id, result.viewingUserId, 'Viewing user should be echoed back.');
    }

    @isTest
    static void testRescheduleAndEndTimeUpdate() {
        ResourceContext context = getActiveResourceContext();

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

        System.assertNotEquals(null, newApptId, 'Appointment should be created.');

        List<FslTechnicianOnlineController.ServiceResourceOption> options =
            FslTechnicianOnlineController.getServiceResourceOptions(wo.Id);
        Test.stopTest();

        System.assertNotEquals(null, newApptId, 'Appointment should be created.');
        System.assert(options != null && !options.isEmpty(), 'Service Resource options should return existing crew members.');
    }

    @isTest
    static void testRequestRescheduleAndMarkQuoteSent() {
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
        FslTechnicianOnlineController.requestRescheduleToResource(wo.Id, resource.Id);
        FslTechnicianOnlineController.markWorkOrderQuoteSent(wo.Id);
        Test.stopTest();
    }
}
