@isTest(SeeAllData=true)
public class FslTechnicianOnlineController_Test {

    /**
     * Locate a Service Resource and User that already exist in the org so we can
     * exercise the queries without inserting synthetic data.
     */
    private static User getExistingTechnicianUser() {
        ServiceResource sr = [
            SELECT Id, RelatedRecordId
            FROM ServiceResource
            WHERE RelatedRecordId != null
            LIMIT 1
        ];

        return [SELECT Id FROM User WHERE Id = :sr.RelatedRecordId LIMIT 1];
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
        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE SchedStartTime != null AND SchedEndTime != null
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
        ServiceAppointment sa = [
            SELECT Id, SchedStartTime, SchedEndTime
            FROM ServiceAppointment
            WHERE SchedStartTime != null AND SchedEndTime != null
            LIMIT 1
            FOR UPDATE
        ];

        ServiceResource resource = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId != null
            LIMIT 1
        ];

        Test.startTest();
        FslTechnicianOnlineController.assignCrewAppointment(sa.Id, resource.Id);
        Test.stopTest();
    }

    @isTest
    static void testCreateAppointmentAndOptions() {
        User techUser = getExistingTechnicianUser();

        WorkOrder wo = [
            SELECT Id, OwnerId, ServiceTerritoryId, Status, Subject, AccountId, Street, City, State, PostalCode, Country
            FROM WorkOrder
            WHERE OwnerId = :techUser.Id
            AND ServiceTerritoryId != null
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

        List<FslTechnicianOnlineController.ServiceResourceOption> options =
            FslTechnicianOnlineController.getServiceResourceOptions(wo.Id);
        Test.stopTest();

        System.assertNotEquals(null, newApptId, 'Appointment should be created.');
        System.assert(options != null && !options.isEmpty(), 'Service Resource options should return existing crew members.');
    }

    @isTest
    static void testRequestRescheduleAndMarkQuoteSent() {
        ServiceResource resource = [
            SELECT Id
            FROM ServiceResource
            WHERE RelatedRecordId != null
            LIMIT 1
        ];

        WorkOrder wo = [
            SELECT Id, OwnerId
            FROM WorkOrder
            WHERE OwnerId = :resource.RelatedRecordId
            LIMIT 1
        ];

        Test.startTest();
        FslTechnicianOnlineController.requestRescheduleToResource(wo.Id, resource.Id);
        FslTechnicianOnlineController.markWorkOrderQuoteSent(wo.Id);
        Test.stopTest();
    }
}
