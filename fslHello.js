import { LightningElement, track } from 'lwc';
import FORM_FACTOR from '@salesforce/client/formFactor';
import { NavigationMixin } from 'lightning/navigation';
import getMyAppointmentsOnline from '@salesforce/apex/FslTechnicianOnlineController.getMyAppointmentsOnline';
import rescheduleAppointment from '@salesforce/apex/FslTechnicianOnlineController.rescheduleAppointment';
import assignCrewAppointment from '@salesforce/apex/FslTechnicianOnlineController.assignCrewAppointment';
import createAppointmentForWorkOrder from '@salesforce/apex/FslTechnicianOnlineController.createAppointmentForWorkOrder';
import updateAppointmentEnd from '@salesforce/apex/FslTechnicianOnlineController.updateAppointmentEnd';
import unassignAppointment from '@salesforce/apex/FslTechnicianOnlineController.unassignAppointment';
import getTerritoryResources from '@salesforce/apex/FslTechnicianOnlineController.getTerritoryResources';
import createEngineerTransferRequest from '@salesforce/apex/FslTechnicianOnlineController.createEngineerTransferRequest';
import acceptEngineerTransferRequest from '@salesforce/apex/FslTechnicianOnlineController.acceptEngineerTransferRequest';
import rejectEngineerTransferRequest from '@salesforce/apex/FslTechnicianOnlineController.rejectEngineerTransferRequest';
import markWorkOrderQuoteSent from '@salesforce/apex/FslTechnicianOnlineController.markWorkOrderQuoteSent';
import updateResourceAbsence from '@salesforce/apex/FslTechnicianOnlineController.updateResourceAbsence';
import deleteResourceAbsence from '@salesforce/apex/FslTechnicianOnlineController.deleteResourceAbsence';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';

export default class FslHello extends NavigationMixin(LightningElement) {
    @track appointments = [];
    @track absences = [];
    @track debugInfo = {};
    @track calendarDays = [];
    @track selectedAppointment = null;
    @track selectedAbsence = null;
    get hasVisibleAppointments() {
        return this.visibleAppointments.length > 0;
    }
    currentUserId = null;
    activeUserId = null;
    viewingUserName = 'Me';
    isManager = false;
    managerTeam = [];
    selectedManagerUserId = null;
    selectedManagerUserName = '';

    // Unscheduled work orders (tray)
    @track unscheduledWorkOrders = [];
    @track transferRequests = [];
    @track submittedTransferRequests = [];
    pullTrayOpen = false;
    pullTrayPeek = false;
    _trayWasExpandedBeforeDrag = false;
    _trayOpenBeforeDrag = false;
    isDesktopFormFactor = FORM_FACTOR === 'Large';
    activeTab = 'list';
    isCalendarTabActive = false;
    lastKnownActiveTab = 'list';

    // Global "now" line state
    showNowLine = false;
    nowLineStyle = '';

    // requestAnimationFrame handle for positioning the "now" line
    _nowLineFrame = null;

    isLoading = false;
    isOffline = false;

    // Center timeline on "today" only when explicitly requested
    _needsCenterOnToday = false;

    // Salesforce user time-zone
    userTimeZoneId = null;
    userTimeZoneShort = null;

    // Calendar config
    daysToShow = 14;
    timelineStartDate;
    weekStartDate;
    calendarStartHour = 0;
    calendarEndHour = 24;

    // Drag and drop state for calendar
    dragMode = null;              // 'event', 'wo', or 'resize'
    draggingEventId = null;
    draggingWorkOrderId = null;
    dragStartDayIndex = null;
    dragCurrentDayIndex = null;
    dragStartLocal = null;
    dragStartClientX = null;
    dragStartClientY = null;
    dragDayWidth = null;
    dragDayBodyTop = null; // screen Y of the top of the day body (for time alignment)
    dragDayBodyHeight = null;
    dragPreviewLocal = null;
    dragStartEndLocal = null;
    dragPreviewDurationHours = null;
    dragDurationHours = null;
    dragHasMoved = false;
    defaultWorkOrderDurationHours = 6;
    showTrayCancelZone = false;
    isHoveringCancelZone = false;
    dragRequiresExplicitConfirmation = false;
    isAwaitingScheduleConfirmation = false;
    pendingSchedulePlacement = null;
    schedulePreviewCardId = null;
    schedulePreviewListMode = null;

    // Reschedule with another tech modal
    isRescheduleModalOpen = false;
    rescheduleOptions = [];
    rescheduleSelection = null;
    rescheduleWorkOrderId = null;
    rescheduleLoading = false;

    // Unschedule confirmation modal
    isUnassignModalOpen = false;
    unassignTarget = null;

    // Transfer request rejection modal
    isRejectModalOpen = false;
    rejectReason = '';
    rejectRequestId = null;

    // Global error capture handlers
    _boundOnGlobalError = null;
    _boundOnUnhandledRejection = null;
    _hasRegisteredErrorHandlers = false;

    // Long press to start drag
    dragLongPressTimer = null;
    dragHoldDelayMs = 600;
    isPressingForDrag = false;
    _pendingDrag = null;          // holds data until long press triggers
    _boundGlobalPointerMove = null;
    _boundGlobalPointerEnd = null;

    // Auto-scroll while dragging near viewport edges
    _autoScrollPoint = null;
    _autoScrollFrame = null;

    // Anchor drag ghost to calendar while awaiting confirmation
    _boundGhostAnchorUpdater = null;

    // Floating ghost under the finger
    // Floating ghost under the finger (clone of the event)
    dragGhostVisible = false;
    dragGhostX = 0;
    dragGhostY = 0;
    dragGhostWidth = 0;
    dragGhostHeight = 0;
    dragGhostTitle = '';
    dragGhostTime = '';
    dragGhostTypeClass = ''; // sfs-event-pm / sfs-event-breakfix / etc

    // 'timeline' or 'week'
    calendarMode = 'timeline';
    isCalendarPanMode = false;
    // list sub-modes: 'my', 'needQuote', 'poRequested', 'quoteSent', 'quotes', 'quoteAttached', 'crew', 'partsReady', 'fulfilling'
    listMode = 'my';

    quoteStatuses = ['Need Quote', 'PO Requested', 'Quote Sent', 'Quote Attached'];

    // My-tab status filter (WorkOrder.Status)
    selectedMyStatus = 'all';

    // For auto-centering timeline
    hasAutoCentered = false;
    todayDayIndex = null;
    _centerTimeout;

    // Detail bottom sheet animation
    isDetailClosing = false;
    _closeTimeout;
    isAbsenceDetailClosing = false;
    _absenceCloseTimeout;

    // ======= GETTERS =======

    get hasAppointments() {
        const apptCount = this.appointments ? this.appointments.length : 0;
        const absenceCount = this.absences ? this.absences.length : 0;
        return apptCount + absenceCount > 0;
    }

    get isMyMode() {
        return this.listMode === 'my';
    }

    get isListTabActive() {
        return this.activeTab === 'list';
    }

    get isManagerTabActive() {
        return this.activeTab === 'manager';
    }

    get isManagerTabVisible() {
        return this.isManager && this.isManagerTabActive;
    }

    get listTabButtonClass() {
        return this.getTabButtonClass('list');
    }

    get calendarTabButtonClass() {
        return this.getTabButtonClass('calendar');
    }

    get managerTabButtonClass() {
        return this.getTabButtonClass('manager');
    }

    getTabButtonClass(tabValue) {
        const baseClass = 'sfs-tab-btn';
        return this.activeTab === tabValue
            ? `${baseClass} sfs-tab-btn_active`
            : baseClass;
    }

    get isDragGhostVisible() {
        return this.dragGhostVisible;
    }

    get isDragging() {
        return !!(this.dragMode || this.isPressingForDrag);
    }

    get calendarDaysWrapperClass() {
        const classes = ['sfs-calendar-days-wrapper'];

        if (this.isCalendarPanMode) {
            classes.push('sfs-calendar-days-wrapper_pan');
        }

        if (this.isDragging) {
            classes.push('sfs-calendar-days-wrapper_dragging');
        }

        return classes.join(' ');
    }

    get calendarDaysClass() {
        return this.isCalendarPanMode
            ? 'sfs-calendar-days sfs-calendar-days_pan'
            : 'sfs-calendar-days';
    }

    get calendarPanButtonClass() {
        const base = 'sfs-calendar-pan-toggle';
        return this.isCalendarPanMode
            ? `${base} sfs-calendar-pan-toggle_active`
            : base;
    }

    get calendarPanStateLabel() {
        return this.isCalendarPanMode ? 'Panning' : 'Drag to move';
    }

    get trayCancelZoneClass() {
        return this.isHoveringCancelZone
            ? 'sfs-tray-cancel-zone sfs-tray-cancel-zone_active'
            : 'sfs-tray-cancel-zone';
    }

    get isViewingAsOther() {
        return (
            this.activeUserId &&
            this.currentUserId &&
            this.activeUserId !== this.currentUserId
        );
    }

    get viewingAsLabel() {
        return this.isViewingAsOther
            ? `Viewing as ${this.viewingUserName}`
            : 'Viewing as yourself';
    }

    get managerOptions() {
        return (this.managerTeam || []).map(m => ({
            label: m.name,
            value: m.userId
        }));
    }

    get managerApplyDisabled() {
        return !this.selectedManagerUserId;
    }

    registerGlobalDragListeners() {
        if (!this._boundGlobalPointerMove) {
            this._boundGlobalPointerMove = this.handleCalendarPointerMove.bind(this);
        }
        if (!this._boundGlobalPointerEnd) {
            this._boundGlobalPointerEnd = this.handleCalendarPointerEnd.bind(this);
        }

        window.addEventListener('mousemove', this._boundGlobalPointerMove);
        window.addEventListener('mouseup', this._boundGlobalPointerEnd);
        window.addEventListener('touchmove', this._boundGlobalPointerMove, {
            passive: false
        });
        window.addEventListener('touchend', this._boundGlobalPointerEnd);
    }

    unregisterGlobalDragListeners() {
        if (this._boundGlobalPointerMove) {
            window.removeEventListener('mousemove', this._boundGlobalPointerMove);
            window.removeEventListener('touchmove', this._boundGlobalPointerMove);
        }

        if (this._boundGlobalPointerEnd) {
            window.removeEventListener('mouseup', this._boundGlobalPointerEnd);
            window.removeEventListener('touchend', this._boundGlobalPointerEnd);
        }
    }

    updateTrayCancelHover(clientX, clientY) {
        if (!this.showTrayCancelZone || this.dragMode !== 'wo') {
            this.isHoveringCancelZone = false;
            return;
        }

        const zone = this.template.querySelector('[data-cancel-zone="true"]');
        if (!zone) {
            this.isHoveringCancelZone = false;
            return;
        }

        const rect = zone.getBoundingClientRect();
        const inside =
            clientX >= rect.left &&
            clientX <= rect.right &&
            clientY >= rect.top &&
            clientY <= rect.bottom;

        this.isHoveringCancelZone = inside;
    }

    isPointInTrayCancelZone(clientX, clientY) {
        const zone = this.template.querySelector('[data-cancel-zone="true"]');
        if (!zone) {
            return false;
        }

        const rect = zone.getBoundingClientRect();
        return (
            clientX >= rect.left &&
            clientX <= rect.right &&
            clientY >= rect.top &&
            clientY <= rect.bottom
        );
    }

    get managerResetDisabled() {
        return !this.isViewingAsOther;
    }

    get rescheduleSubmitDisabled() {
        return this.rescheduleLoading || !this.rescheduleSelection;
    }

    get absenceDebugSummary() {
        const absCount = this.absences ? this.absences.length : 0;
        const ids =
            this.debugInfo && Array.isArray(this.debugInfo.absenceResourceIds)
                ? this.debugInfo.absenceResourceIds
                : [];
        const resourceLabel = ids.length ? ids.join(', ') : 'none';
        return `Absence fetch: ${absCount} record(s); resources: ${resourceLabel}`;
    }

    get absenceDebugRows() {
        const absences = this.absences || [];
        return absences.map(abs => {
            const startLocal = abs.start
                ? this.convertUtcToUserLocal(abs.start)
                : null;
            const endLocal = abs.endTime
                ? this.convertUtcToUserLocal(abs.endTime)
                : null;
            const rangeLabel =
                startLocal && endLocal
                    ? this.formatTimeRange(startLocal, endLocal)
                    : 'No time available';

            return {
                key: abs.absenceId || abs.subject || rangeLabel,
                subject: abs.subject || 'Absence',
                resourceId: abs.resourceId,
                range: rangeLabel
            };
        });
    }

    get hasAbsenceDebugRows() {
        return this.absenceDebugRows.length > 0;
    }

    // Position + size of the floating event
    get dragGhostStyle() {
        return `top:${this.dragGhostY}px;left:${this.dragGhostX}px;width:${this.dragGhostWidth}px;height:${this.dragGhostHeight}px;transform:translateX(-50%);`;
    }

    get dragGhostWrapperClass() {
        return this.showDragConfirmActions
            ? 'sfs-drag-ghost sfs-drag-ghost_interactive'
            : 'sfs-drag-ghost';
    }


    // Classes for the inner event block
    get dragGhostClass() {
        const base = 'sfs-calendar-event sfs-calendar-event_ghost';
        return this.dragGhostTypeClass
            ? `${base} ${this.dragGhostTypeClass}`
            : base;
    }

    get dragHelperText() {
        if (this.showDragConfirmActions) {
            return 'Tap ✓ to schedule or ✕ to cancel';
        }

        return '';
    }

    get showDragConfirmActions() {
        return (
            (this.dragRequiresExplicitConfirmation && this.dragGhostVisible) ||
            (this.isAwaitingScheduleConfirmation && !!this.pendingSchedulePlacement)
        );
    }


    get isCrewCountUrgent() {
        if (!this.appointments || !this.appointments.length) {
            return false;
        }

        const now = new Date();
        const in48 = new Date(now.getTime() + 48 * 60 * 60 * 1000);

        return this.appointments.some(a => {
            if (!a.isCrewAssignment || a.isMyAssignment || !a.schedStart) {
                return false;
            }
            const start = new Date(a.schedStart);
            return start >= now && start <= in48;
        });
    }

    get myCount() {
        if (!this.appointments) return 0;

        const anyFlagged = this.appointments.some(
            a => a.isMyAssignment || a.isCrewAssignment
        );

        if (anyFlagged) {
            return this.appointments.filter(
                a => a.isMyAssignment && !a.isCrewAssignment
            ).length;
        }

        return this.appointments.length;
    }

    get crewCount() {
        if (!this.appointments) return 0;
        return this.appointments.filter(a => a.isCrewAssignment).length;
    }

    get transferRequestCount() {
        return (this.transferRequests && Array.isArray(this.transferRequests))
            ? this.transferRequests.length
            : 0;
    }

    get submittedTransferRequestCount() {
        return this.submittedTransferRequests.length;
    }

    get partsReadyCount() {
        if (!this.appointments) return 0;
        return this.appointments.filter(a => a.allPartsEnRoute).length;
    }

    get fulfillingCount() {
        if (!this.appointments) return 0;
        return this.appointments.filter(
            a => a.somePartsEnRoute && !a.allPartsEnRoute
        ).length;
    }

    // Status options for My tab filter (WorkOrder.Status)
    get myStatusOptions() {
        const statuses = new Set();

        this.appointments.forEach(a => {
            const anyFlagged = a.isMyAssignment || a.isCrewAssignment;
            let isMy = false;

            if (anyFlagged) {
                isMy = a.isMyAssignment && !a.isCrewAssignment;
            } else {
                isMy = true;
            }

            if (isMy && a.workOrderStatus) {
                statuses.add(a.workOrderStatus);
            }
        });

        const statusArray = Array.from(statuses).sort();
        const options = statusArray.map(s => ({
            label: s,
            value: s
        }));

        return [{ label: 'All', value: 'all' }, ...options];
    }

    get ownedAppointments() {
        if (!this.appointments) return [];

        const anyFlagged = this.appointments.some(
            a => a.isMyAssignment || a.isCrewAssignment
        );

        if (anyFlagged) {
            return this.appointments.filter(
                a => a.isMyAssignment && !a.isCrewAssignment
            );
        }

        return this.appointments;
    }

    get quotesCount() {
        return (
            this.ownedAppointments.filter(appt =>
                this.isQuoteStatus(appt.workOrderStatus)
            ).length + this.quoteWorkOrders.length
        );
    }

    get needQuoteCount() {
        return (
            this.ownedAppointments.filter(appt =>
                appt.workOrderStatus === 'Need Quote'
            ).length +
            this.quoteWorkOrders.filter(wo => wo.workOrderStatus === 'Need Quote')
                .length
        );
    }

    get poRequestedCount() {
        return (
            this.ownedAppointments.filter(appt =>
                appt.workOrderStatus === 'PO Requested'
            ).length +
            this.quoteWorkOrders.filter(
                wo => wo.workOrderStatus === 'PO Requested'
            ).length
        );
    }

    get quoteSentCount() {
        return (
            this.ownedAppointments.filter(appt =>
                appt.workOrderStatus === 'Quote Sent'
            ).length +
            this.quoteWorkOrders.filter(
                wo => wo.workOrderStatus === 'Quote Sent'
            ).length
        );
    }

    get quoteAttachedCount() {
        return (
            this.ownedAppointments.filter(appt =>
                this.isQuoteAttachedAppointment(appt)
            ).length +
            this.quoteWorkOrders.filter(wo =>
                this.isQuoteAttachedAppointment(wo)
            ).length
        );
    }

    get visibleAppointments() {
        if (!this.appointments) return [];

        let baseList = [];

        const ownedAppointments = this.ownedAppointments.map(appt => ({
            ...appt,
            cardId: appt.appointmentId,
            hasAppointment: true
        }));

        const quoteWorkOrders = this.quoteWorkOrders;
        const unscheduledWorkOrders = this.unscheduledListItems;

        switch (this.listMode) {
            case 'crew':
                baseList = ownedAppointments.filter(a => a.isCrewAssignment);
                break;

            case 'partsReady':
                baseList = ownedAppointments.filter(a => a.allPartsEnRoute);
                break;

            case 'fulfilling':
                baseList = ownedAppointments.filter(
                    a => a.somePartsEnRoute && !a.allPartsEnRoute
                );
                break;

            case 'quotes':
                baseList = ownedAppointments
                    .filter(appt => this.isQuoteStatus(appt.workOrderStatus))
                    .concat(quoteWorkOrders);
                break;

            case 'needQuote':
                baseList = ownedAppointments
                    .filter(appt => appt.workOrderStatus === 'Need Quote')
                    .concat(
                        quoteWorkOrders.filter(
                            wo => wo.workOrderStatus === 'Need Quote'
                        )
                    );
                break;

            case 'poRequested':
                baseList = ownedAppointments
                    .filter(appt => appt.workOrderStatus === 'PO Requested')
                    .concat(
                        quoteWorkOrders.filter(
                            wo => wo.workOrderStatus === 'PO Requested'
                        )
                    );
                break;

            case 'quoteSent':
                baseList = ownedAppointments
                    .filter(appt => appt.workOrderStatus === 'Quote Sent')
                    .concat(
                        quoteWorkOrders.filter(
                            wo => wo.workOrderStatus === 'Quote Sent'
                        )
                    );
                break;

            case 'quoteAttached':
                baseList = ownedAppointments
                    .filter(appt => this.isQuoteAttachedAppointment(appt))
                    .concat(
                        quoteWorkOrders.filter(
                            wo => this.isQuoteAttachedAppointment(wo)
                        )
                    );
                break;

            case 'unscheduled':
                baseList = unscheduledWorkOrders;
                break;

            case 'my':
            default: {
                baseList = ownedAppointments;
                break;
            }
        }

        if (
            this.listMode === 'my' &&
            this.selectedMyStatus &&
            this.selectedMyStatus !== 'all'
        ) {
            baseList = baseList.filter(
                a => a.workOrderStatus === this.selectedMyStatus
            );
        }

        const allowScheduleOnCalendar =
            this.listMode === 'unscheduled' || this.listMode === 'partsReady';

        return baseList.map(item => ({
            ...item,
            showScheduleOnCalendar: allowScheduleOnCalendar
        }));
    }

    isQuoteStatus(status) {
        return this.quoteStatuses.includes(status);
    }

    get quoteWorkOrders() {
        if (!this.unscheduledWorkOrders) {
            return [];
        }

        return this.unscheduledWorkOrders
            .filter(wo => this.isQuoteStatus(wo.status))
            .map(wo => ({
                cardId: `wo-${wo.workOrderId}`,
                appointmentId: null,
                hasAppointment: false,
                workOrderId: wo.workOrderId,
                workOrderStatus: wo.status,
                workOrderNumber: wo.workOrderNumber,
                workOrderSubject: wo.subject,
                accountName: wo.accountName,
                fullAddress: wo.fullAddress,
                opportunityRecordType: wo.opportunityRecordType,
                quoteAttachmentUrl: wo.quoteAttachmentDownloadUrl || null,
                quoteAttachmentDocumentId:
                    wo.quoteAttachmentDocumentId || null,
                hasQuoteAttachment: Boolean(
                    wo.hasQuoteAttachment ||
                        wo.quoteAttachmentDownloadUrl ||
                        wo.quoteAttachmentDocumentId
                ),
                showQuoteActions:
                    wo.status === 'Quote Sent' ||
                    this.isQuoteAttachedAppointment(wo),
                showMarkQuoteSentAction: this.isQuoteAttachedAppointment(wo),
                isExpanded: false,
                workTypeName: 'Work Order',
                workTypeClass: 'sfs-worktype'
            }));
    }

    get unscheduledListItems() {
        if (!this.unscheduledWorkOrders) {
            return [];
        }

        return this.unscheduledWorkOrders.map(wo => {
            const typeClass = this.getEventTypeClass(wo.workTypeName);
            return {
                ...wo,
                cardId: wo.cardId || `wo-${wo.workOrderId}`,
                workOrderId: wo.workOrderId,
                workOrderSubject: wo.workOrderSubject || wo.subject,
                workOrderStatus: wo.workOrderStatus || wo.status,
                workOrderNumber: wo.workOrderNumber,
                workTypeName: wo.workTypeName,
                workTypeClass: `sfs-worktype ${typeClass || ''}`.trim(),
                opportunityRecordType: wo.opportunityRecordType,
                quoteAttachmentUrl: wo.quoteAttachmentDownloadUrl || null,
                quoteAttachmentDocumentId: wo.quoteAttachmentDocumentId || null,
                hasQuoteAttachment: Boolean(
                    wo.hasQuoteAttachment ||
                        wo.quoteAttachmentDownloadUrl ||
                        wo.quoteAttachmentDocumentId
                ),
                showQuoteActions:
                    wo.status === 'Quote Sent' ||
                    this.isQuoteAttachedAppointment(wo),
                showMarkQuoteSentAction: this.isQuoteAttachedAppointment(wo),
                hasAppointment: false
            };
        });
    }

    isQuoteAttachedAppointment(appt) {
        const status = (
            appt.workOrderStatus ||
            appt.status ||
            ''
        ).toLowerCase();
        const hasAttachment = appt.hasQuoteAttachment || Boolean(
            appt.quoteAttachmentUrl ||
                appt.quoteAttachmentDownloadUrl ||
                appt.quoteAttachmentDocumentId
        );

        return status.startsWith('quote attached') ||
            (status === 'need quote' && hasAttachment);
    }

    findAppointmentByCardId(cardId) {
        if (!cardId) {
            return null;
        }

        const visible = (this.visibleAppointments || []).find(
            a => a.cardId === cardId
        );
        if (visible) {
            return visible;
        }

        const apptMatch = (this.appointments || []).find(
            a => a.cardId === cardId || a.appointmentId === cardId
        );
        if (apptMatch) {
            return apptMatch;
        }

        return (this.quoteWorkOrders || []).find(wo => wo.cardId === cardId);
    }

    normalizeWorkOrderDetail(workOrder) {
        if (!workOrder) {
            return null;
        }

        const typeClass = this.getEventTypeClass(workOrder.workTypeName);

        return {
            ...workOrder,
            appointmentId: null,
            workOrderId: workOrder.workOrderId,
            workOrderSubject: workOrder.workOrderSubject || workOrder.subject,
            workOrderStatus: workOrder.status,
            workOrderNumber: workOrder.workOrderNumber,
            workTypeName: workOrder.workTypeName,
            workTypeClass: `sfs-worktype ${typeClass || ''}`.trim(),
            schedStart: null,
            schedEnd: null,
            newStart: null,
            disableSave: true,
            isExpanded: false,
            hasAppointment: false,
            cardId: workOrder.cardId || `wo-${workOrder.workOrderId}`
        };
    }

    findAbsenceById(absenceId) {
        if (!absenceId || !this.absences) {
            return null;
        }

        return this.absences.find(a => a.absenceId === absenceId) || null;
    }

    get hasSelectedAppointment() {
        return this.selectedAppointment !== null;
    }

    get hasSelectedAbsence() {
        return this.selectedAbsence !== null;
    }

    get calendarHourLabels() {
        const labels = [];
        for (let h = this.calendarStartHour; h < this.calendarEndHour; h++) {
            const dt = new Date(2020, 0, 1, h, 0, 0, 0);
            labels.push(
                dt.toLocaleTimeString([], {
                    hour: 'numeric'
                })
            );
        }
        return labels;
    }

    get isCrewMode() {
        return this.listMode === 'crew';
    }

    get isTransferMode() {
        return this.listMode === 'transferRequests';
    }

    get listModeOptions() {
        return [
            this.buildListModeOption(
                'unscheduled',
                'Unscheduled',
                this.unscheduledCount
            ),
            this.buildListModeOption('my', 'Scheduled', this.myCount),
            this.buildListModeOption(
                'transferRequests',
                'Transfer Requests',
                this.transferRequestCount
            ),
            this.buildListModeOption('crew', 'Crew Pool', this.crewCount),
            this.buildListModeOption(
                'needQuote',
                'Quote Needed',
                this.needQuoteCount
            ),
            this.buildListModeOption(
                'poRequested',
                'PO Requested',
                this.poRequestedCount
            ),
            this.buildListModeOption(
                'quoteAttached',
                'Quote Attached',
                this.quoteAttachedCount
            ),
            this.buildListModeOption(
                'quoteSent',
                'Quote Sent',
                this.quoteSentCount
            ),
            this.buildListModeOption(
                'partsReady',
                'Parts Ready',
                this.partsReadyCount
            ),
            this.buildListModeOption(
                'fulfilling',
                'Currently Fulfilling',
                this.fulfillingCount
            )
        ];
    }

    get listModeChips() {
        return this.listModeOptions
            .filter(opt => opt.count > 0 || opt.value === this.listMode)
            .map(opt => ({
                value: opt.value,
                label: opt.label,
                className: this.getListModeChipClass(opt.value, opt.count),
                isActive: opt.value === this.listMode
            }));
    }

    buildListModeOption(value, label, count) {
        return {
            value,
            label: `${label} (${count})`,
            count
        };
    }

    getListModeChipClass(modeValue, count) {
        let classes = 'sfs-mode-chip';

        if (modeValue === this.listMode) {
            classes += ' sfs-mode-chip_active';
        }
        if (modeValue === 'transferRequests' && count > 0) {
            classes += ' sfs-mode-chip_alert';
        }
        if (modeValue === 'crew' && this.isCrewCountUrgent) {
            classes += ' sfs-mode-chip_warning';
        }
        return classes;
    }

    get isTimelineMode() {
        return this.calendarMode === 'timeline';
    }

    get isWeekMode() {
        return this.calendarMode === 'week';
    }

    get timelineModeClass() {
        return this.calendarMode === 'timeline'
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get weekModeClass() {
        return this.calendarMode === 'week'
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    // Calendar -> weeks
    get calendarWeeks() {
        const weeks = [];
        let current = [];
        let weekIndex = 0;

        this.calendarDays.forEach((day, index) => {
            current.push(day);
            if ((index + 1) % 7 === 0) {
                weeks.push({
                    key: `week-${weekIndex}`,
                    days: current
                });
                weekIndex += 1;
                current = [];
            }
        });

        if (current.length) {
            weeks.push({
                key: `week-${weekIndex}`,
                days: current
            });
        }

        return weeks;
    }

    get calendarHeaderDays() {
        const weeks = this.calendarWeeks;
        return weeks && weeks.length ? weeks[0].days : [];
    }

    get calendarRangeLabel() {
        if (!this.calendarDays || this.calendarDays.length === 0) return '';

        const first = this.calendarDays[0].date;
        const last = this.calendarDays[this.calendarDays.length - 1].date;

        const sameMonth =
            first.getMonth() === last.getMonth() &&
            first.getFullYear() === last.getFullYear();

        const optsStart = {
            month: 'short',
            day: 'numeric',
            timeZone: this.userTimeZoneId || undefined
        };
        const optsEnd = sameMonth
            ? {
                  day: 'numeric',
                  year: 'numeric',
                  timeZone: this.userTimeZoneId || undefined
              }
            : {
                  month: 'short',
                  day: 'numeric',
                  year: 'numeric',
                  timeZone: this.userTimeZoneId || undefined
              };

        return (
            first.toLocaleDateString([], optsStart) +
            ' – ' +
            last.toLocaleDateString([], optsEnd)
        );
    }

    get detailCardClass() {
        return this.isDetailClosing
            ? 'sfs-detail-card sfs-detail-card_closing'
            : 'sfs-detail-card sfs-detail-card_open';
    }

    get absenceDetailCardClass() {
        return this.isAbsenceDetailClosing
            ? 'sfs-detail-card sfs-detail-card_closing'
            : 'sfs-detail-card sfs-detail-card_open';
    }

    // Tray helpers
    get trayContainerClass() {
        const classes = ['sfs-tray'];

        if (this.pullTrayOpen) {
            classes.push('sfs-tray_open');
        }

        if (this.pullTrayPeek) {
            classes.push('sfs-tray_peek');
        }

        if (this.pullTrayOpen && (this.dragMode || this.isPressingForDrag)) {
            classes.push('sfs-tray_dragging');
        }

        return classes.join(' ');
    }

    get showTrayPeekToggle() {
        return this.pullTrayOpen && this.shouldUseCompactTray();
    }

    get trayPeekToggleLabel() {
        return this.pullTrayPeek ? 'Expand full tray' : 'Peek to view';
    }

    get trayPeekHelperText() {
        return this.pullTrayPeek
            ? 'Peek view keeps the calendar visible.'
            : 'Peek to keep the calendar in view while browsing.';
    }

    get unscheduledCount() {
        return this.unscheduledWorkOrders
            ? this.unscheduledWorkOrders.length
            : 0;
    }

    get firstTimeWorkOrders() {
        const workOrders = this.unscheduledWorkOrders || [];
        return workOrders.filter(wo => {
            const recordType = (wo.recordTypeName || '').toLowerCase();
            const saCount = wo.serviceAppointmentCount || 0;
            const unscheduledCount = wo.unscheduledServiceAppointmentCount || 0;

            if (recordType !== 'fsl work order') {
                return false;
            }

            if (saCount === 0) {
                return true;
            }

            return saCount === 1 && unscheduledCount === 1;
        });
        return workOrders.filter(
            wo => wo.serviceAppointmentCount === 1 && !wo.needsReturnVisitScheduling
        );
    }

    get returnVisitWorkOrders() {
        const workOrders = this.unscheduledWorkOrders || [];
        return workOrders.filter(wo => {
            const recordType = (wo.recordTypeName || '').toLowerCase();
            const saCount = wo.serviceAppointmentCount || 0;
            const unscheduledCount = wo.unscheduledServiceAppointmentCount || 0;
            const isFirstTimeCandidate =
                saCount === 0 || (saCount === 1 && unscheduledCount === 1);

            if (recordType !== 'fsl work order' || saCount === 0) {
                return false;
            }

            if (isFirstTimeCandidate) {
                return false;
            }

            const allReturnRequired =
                wo.allServiceAppointmentsReturnRequired === true;
            const hasUnscheduledNonReturn =
                wo.hasUnscheduledNonReturnAppointment === true;

            return (allReturnRequired && unscheduledCount === 0) ||
                hasUnscheduledNonReturn;
        });
        return workOrders.filter(wo => wo.needsReturnVisitScheduling);
    }

    get firstTimeCount() {
        return this.firstTimeWorkOrders.length;
    }

    get returnVisitCount() {
        return this.returnVisitWorkOrders.length;
    }

    get hasUnscheduled() {
        return this.unscheduledCount > 0;
    }

    /**
     * Friendly text for the tray handle
     * e.g. "3 work orders need scheduling" or "No work orders need scheduling"
     */
    get unscheduledLabel() {
        const count = this.unscheduledCount;

        if (count === 0) {
            return 'No work orders need scheduling';
        }
        if (count === 1) {
            return '1 work order needs scheduling';
        }
        return `${count} work orders need scheduling`;
    }

    get pullUpTrayCriteria() {
        const criteria = [
            {
                key: 'calendarTab',
                label: 'Calendar tab is active',
                met: this.isCalendarTabActive
            },
            {
                key: 'dataLoaded',
                label: 'Data finished loading',
                met: !this.isLoading
            },
            {
                key: 'unscheduledLoaded',
                label: 'Unscheduled work orders loaded',
                met: Array.isArray(this.unscheduledWorkOrders)
            },
            {
                key: 'online',
                label: 'Online (required for scheduling)',
                met: !this.isOffline
            }
        ];

        return criteria.map(item => ({
            ...item,
            statusText: item.met ? 'Met' : 'Missing',
            statusSymbol: item.met ? '✔' : '⚠',
            itemClass: item.met
                ? 'sfs-tray-checklist__item sfs-tray-checklist__item_met'
                : 'sfs-tray-checklist__item sfs-tray-checklist__item_missing'
        }));
    }

    get isTrayReady() {
        return this.pullUpTrayCriteria.every(item => item.met);
    }

    get trayReadinessSummary() {
        return this.isTrayReady
            ? 'All criteria are satisfied. The pull-up bar should appear below.'
            : 'One or more criteria are missing. Resolve the items below to render the pull-up bar.';
    }

    get trayStatusClass() {
        return this.isTrayReady
            ? 'sfs-tray-debug__badge sfs-tray-debug__badge_ready'
            : 'sfs-tray-debug__badge sfs-tray-debug__badge_blocked';
    }

    get trayStatusLabel() {
        return this.isTrayReady ? 'Ready' : 'Blocked';
    }

    get shouldRenderTray() {
        return this.isTrayReady;
    }

    updateActiveTabState(explicitValue, options = {}) {
        const { suppressCalendarToday = false } = options;
        let resolvedTabValue = explicitValue;

        if (resolvedTabValue === undefined || resolvedTabValue === null) {
            resolvedTabValue =
                this.activeTab ||
                this.lastKnownActiveTab ||
                'list';
        }

        if (resolvedTabValue === 'manager' && !this.isManager) {
            resolvedTabValue = 'list';
        }

        if (resolvedTabValue) {
            this.activeTab = resolvedTabValue;
            this.lastKnownActiveTab = resolvedTabValue;
        }

        const isCalendarActive = resolvedTabValue === 'calendar';
        const wasCalendarActive = this.isCalendarTabActive;

        if (isCalendarActive !== this.isCalendarTabActive) {
            this.isCalendarTabActive = isCalendarActive;

            if (!isCalendarActive) {
                this.pullTrayOpen = false;
            } else if (!wasCalendarActive && !suppressCalendarToday) {
                // Ensure the calendar recenters on today whenever the user
                // switches into the calendar tab (keyboard, click, or
                // programmatic activation).
                this.handleCalendarToday();
            }
        }
    }

    // ======= LIFECYCLE =======

    connectedCallback() {
        try {
            this.registerGlobalErrorHandlers();
            this.checkOnline();
            this.handleCalendarToday();
            if (!this.isOffline) {
                this.loadAppointments();
            }
        } catch (error) {
            this.captureError(error, 'connectedCallback');
        }
    }

    disconnectedCallback() {
        this.unregisterGlobalErrorHandlers();
    }

    renderedCallback() {
        try {
            this.updateActiveTabState();

            if (
                this.isTimelineMode &&
                this.isCalendarTabActive &&
                this._needsCenterOnToday &&
                this.calendarDays &&
                this.calendarDays.length > 0
            ) {
                this.centerTimelineOnTodayColumn();
            }

            this.scheduleNowLinePositionUpdate();
        } catch (error) {
            this.captureError(error, 'renderedCallback');
        }
    }

    errorCallback(error, stack) {
        this.captureError(error, 'errorCallback');

        if (!stack) {
            return;
        }

        const currentLastError = (this.debugInfo && this.debugInfo.lastError) || {};

        this.debugInfo = Object.assign({}, this.debugInfo, {
            lastError: Object.assign({}, currentLastError, {
                stack: stack
            })
        });
        const lastError = (this.debugInfo && this.debugInfo.lastError) || {};

        this.debugInfo = Object.assign({}, this.debugInfo, {
            lastError: Object.assign({}, lastError, {
                stack: stack
            })
        });
        lastError = this.debugInfo?.lastError || {};

        this.debugInfo = {
            ...this.debugInfo,
            lastError: {
                ...lastError,
                stack
            }
        };
    }

    // ======= ONLINE CHECK =======

    checkOnline() {
        if (typeof navigator !== 'undefined' && navigator.onLine === false) {
            this.isOffline = true;
        } else {
            this.isOffline = false;
        }
    }

    // ======= CALENDAR RANGE CONTROL =======

    centerCalendarOnToday() {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const timelineStart = new Date(today);
        timelineStart.setDate(today.getDate() - 7);
        this.timelineStartDate = timelineStart;

        const dow = today.getDay();
        const weekStart = new Date(today);
        weekStart.setDate(today.getDate() - dow);
        this.weekStartDate = weekStart;

        this.hasAutoCentered = false;
        this.buildCalendarModel();
    }

    centerTimelineOnTodayColumn() {
        const wrapper = this.template.querySelector(
            '.sfs-calendar-days-wrapper'
        );
        const todayCol = this.template.querySelector(
            '.sfs-calendar-day_today'
        );
        const daysContainer = this.template.querySelector(
            '.sfs-calendar-days'
        );

        if (!wrapper || !daysContainer || !todayCol) {
            this._needsCenterOnToday = false;
            return;
        }

        const wrapperWidth = wrapper.clientWidth;
        const todayCenter =
            todayCol.offsetLeft + todayCol.offsetWidth / 2;

        const targetScrollLeft = Math.max(
            0,
            todayCenter - wrapperWidth / 2
        );

        wrapper.scrollLeft = targetScrollLeft;
        this.centerTimelineOnElevenAmLine(todayCol);
        this._needsCenterOnToday = false;
    }

    centerTimelineOnElevenAmLine(todayCol) {
        if (typeof window === 'undefined') {
            return;
        }

        const dayBody =
            todayCol?.querySelector('.sfs-calendar-day-body') ||
            this.template.querySelector('.sfs-calendar-day-body');

        if (!dayBody) {
            return;
        }

        const bodyRect = dayBody.getBoundingClientRect();
        const elevenAmOffset = (11 / 24) * bodyRect.height;
        const targetScrollTop =
            bodyRect.top + window.scrollY + elevenAmOffset - window.innerHeight / 2;

        window.scrollTo({
            top: Math.max(targetScrollTop, 0)
        });
    }

    shiftCalendar(offsetDays) {
        if (!this.timelineStartDate && !this.weekStartDate) {
            this.centerCalendarOnToday();
            return;
        }

        if (this.isTimelineMode) {
            const start = new Date(this.timelineStartDate || new Date());
            start.setDate(start.getDate() + offsetDays);
            this.timelineStartDate = start;
        } else {
            const start = new Date(this.weekStartDate || new Date());
            start.setDate(start.getDate() + offsetDays);
            this.weekStartDate = start;
        }

        this.hasAutoCentered = false;
        this.buildCalendarModel();
    }

    getClientPoint(event) {
        if (!event) {
            return null;
        }

        const touch =
            (event.touches && event.touches[0]) ||
            (event.changedTouches && event.changedTouches[0]);

        if (touch) {
            return { clientX: touch.clientX, clientY: touch.clientY };
        }

        if (
            typeof event.clientX === 'number' &&
            typeof event.clientY === 'number'
        ) {
            return { clientX: event.clientX, clientY: event.clientY };
        }

        return null;
    }

    // ======= DRAG HANDLERS (events) =======

    handleEventDragStart(event) {
        // Allow action buttons inside an event (e.g., remove/unschedule) to work on touch
        // devices without being intercepted by drag handlers.
        if (event.target && event.target.closest('.sfs-unassign-btn')) {
            return;
        }

        if (this.isCalendarPanMode) {
            return;
        }

        // Do not start another drag if one is already running
        if (this.dragMode || this.isPressingForDrag) {
            return;
        }

        if (event.currentTarget.dataset.kind === 'absence') {
            return;
        }

        const id = event.currentTarget.dataset.id;
        const dayIndexStr = event.currentTarget.dataset.dayIndex;
        if (!id || dayIndexStr === undefined) {
            return;
        }

        const appt = this.appointments.find(a => a.appointmentId === id);
        if (!appt || !appt.schedStart) {
            return;
        }

        const clientPoint = this.getClientPoint(event);
        if (!clientPoint) {
            return;
        }

        const dayIndex = parseInt(dayIndexStr, 10);

        const localStart = this.convertUtcToUserLocal(appt.schedStart);

        const dayBodyEl = event.currentTarget.closest('.sfs-calendar-day-body');
        const dayEl = event.currentTarget.closest('.sfs-calendar-day');

        if (!dayBodyEl || !dayEl) {
            return;
        }

        const bodyRect = dayBodyEl.getBoundingClientRect();

        const pending = {
            type: 'event',
            id,
            dayIndex,
            localStart,
            clientX: clientPoint.clientX,
            clientY: clientPoint.clientY,
            dayBodyHeight: bodyRect.height || dayBodyEl.clientHeight || 1,
            dayBodyTop: bodyRect.top,
            dayWidth: dayEl.clientWidth || 1,
            title: appt.workOrderSubject || 'Appointment'
        };


        this.isPressingForDrag = true;
        this._pendingDrag = pending;

        // Long press threshold: require a deliberate press-and-hold to move
        this.clearLongPressTimer();
        this.dragLongPressTimer = this.safeSetTimeout(() => {
            this.beginDragFromPending();
        }, this.dragHoldDelayMs);

        // On touch devices, allow the synthetic click event to fire so a quick tap
        // opens the info drawer. Prevent default only for mouse/pen interactions to
        // avoid suppressing the click on mobile while still stopping accidental
        // text selection when dragging with a mouse.
        const isTouchStart = event.type === 'touchstart';
        if (!isTouchStart) {
            event.preventDefault();
            event.stopPropagation();
        }
    }

    handleEventPressEnd() {
        // If we were only waiting for a long press and never started a drag, cancel it
        if (this.isPressingForDrag && !this.dragMode) {
            this.isPressingForDrag = false;
            this._pendingDrag = null;
            this.clearLongPressTimer();
        }
    }

    handleEventResizeStart(event) {
        if (!event) {
            return;
        }

        if (this.isCalendarPanMode) {
            return;
        }

        if (this.dragMode) {
            return;
        }

        if (event.currentTarget.dataset.kind === 'absence') {
            return;
        }

        const id = event.currentTarget.dataset.id;
        const dayIndexStr = event.currentTarget.dataset.dayIndex;
        if (!id || !dayIndexStr) {
            return;
        }

        const dayIndex = parseInt(dayIndexStr, 10);
        const appt = this.appointments.find(a => a.appointmentId === id);
        if (!appt || !appt.schedStart) {
            return;
        }

        const clientPoint = this.getClientPoint(event);
        if (!clientPoint) {
            return;
        }

        const dayBodyEl = event.currentTarget.closest('.sfs-calendar-day-body');
        const dayEl = event.currentTarget.closest('.sfs-calendar-day');
        const eventEl = event.currentTarget.closest('.sfs-calendar-event');

        if (!dayBodyEl || !dayEl || !eventEl) {
            return;
        }

        const bodyRect = dayBodyEl.getBoundingClientRect();
        const dayRect = dayEl.getBoundingClientRect();
        const eventRect = eventEl.getBoundingClientRect();

        const startLocal = this.convertUtcToUserLocal(appt.schedStart);
        const endLocal = appt.schedEnd
            ? this.convertUtcToUserLocal(appt.schedEnd)
            : new Date(startLocal.getTime() + 60 * 60 * 1000);

        const durationHours = this.computeDurationHours(startLocal, endLocal);

        this.dragMode = 'resize';
        this.draggingEventId = id;
        this.draggingWorkOrderId = null;
        this.dragStartDayIndex = dayIndex;
        this.dragCurrentDayIndex = dayIndex;
        this.dragStartLocal = startLocal;
        this.dragStartEndLocal = endLocal;
        this.dragDurationHours = durationHours;
        this.dragPreviewDurationHours = durationHours;
        this.dragStartClientX = clientPoint.clientX;
        this.dragStartClientY = clientPoint.clientY;
        this.dragDayBodyHeight = bodyRect.height || dayBodyEl.clientHeight || 1;
        this.dragDayBodyTop = bodyRect.top;
        this.dragDayWidth = dayRect.width || 1;
        this.dragHasMoved = false;
        this.dragPreviewLocal = new Date(startLocal);

        const totalHours = this.calendarEndHour - this.calendarStartHour;
        const startHourFraction =
            startLocal.getHours() + startLocal.getMinutes() / 60;
        const yWithinBody =
            ((startHourFraction - this.calendarStartHour) / totalHours) *
            this.dragDayBodyHeight;

        const ghostHeight =
            (durationHours / totalHours) * this.dragDayBodyHeight;
        const ghostX = dayRect.left + dayRect.width / 2;
        const timeLabel = this.formatTimeRange(startLocal, endLocal);
        const typeClass = this.getEventTypeClass(appt.workTypeName);

        this.showDragGhost(
            ghostX,
            this.dragDayBodyTop + yWithinBody,
            appt.workOrderSubject || 'Appointment',
            timeLabel,
            typeClass,
            eventRect.width,
            ghostHeight
        );

        this.updateSelectedEventStyles();
        this.showTrayCancelZone = false;
        this.registerGlobalDragListeners();

        if (typeof event.stopPropagation === 'function') {
            event.stopPropagation();
        }

        if (typeof event.preventDefault === 'function') {
            event.preventDefault();
        }
    }



    // ======= DRAG HANDLERS (tray -> calendar) =======
    handleTrayCardDragStart(event) {
        if (this.dragMode || this.isPressingForDrag) {
            return;
        }

        if (!event || !event.currentTarget) {
            this.resetDragState();
            return;
        }

        const card = event.currentTarget;
        const workOrderId = card.dataset && card.dataset.woid;
        if (!workOrderId) {
            return;
        }

        const clientPoint = this.getClientPoint(event);
        if (!clientPoint) {
            return;
        }

        // Use first day column/body to measure width and height
        const dayEl = this.template.querySelector('.sfs-calendar-day');
        const dayBodyEl = this.template.querySelector('.sfs-calendar-day-body');

        if (!dayEl || !dayBodyEl) {
            return;
        }

        if (this.pullTrayOpen && this.shouldUseCompactTray()) {
            this._trayWasExpandedBeforeDrag = !this.pullTrayPeek;
            this.pullTrayPeek = true;
        }

        const startIndex =
            this.todayDayIndex != null ? this.todayDayIndex : 0;

        const bodyRect = dayBodyEl.getBoundingClientRect();

        const pending = {
            type: 'wo',
            workOrderId,
            dayIndex: startIndex,
            clientX: clientPoint.clientX,
            clientY: clientPoint.clientY,
            dayBodyHeight: bodyRect.height || dayBodyEl.clientHeight || 1,
            dayBodyTop: bodyRect.top,
            dayWidth: dayEl.clientWidth || 1,
            title: card.querySelector('.sfs-tray-card-title')
                ? card.querySelector('.sfs-tray-card-title').textContent
                : 'New appointment'
        };

        this.isPressingForDrag = true;
        this._pendingDrag = pending;

        this.clearLongPressTimer();
        this.dragLongPressTimer = this.safeSetTimeout(() => {
            this.beginDragFromPending();
        }, this.dragHoldDelayMs);

        event.preventDefault();
        event.stopPropagation();
    }

    clearLongPressTimer() {
        if (this.dragLongPressTimer) {
            this.safeClearTimeout(this.dragLongPressTimer);
            this.dragLongPressTimer = null;
        }
    }

    compactTrayForDrag() {
        if (!this.pullTrayOpen) {
            this._trayOpenBeforeDrag = false;
            return;
        }

        this._trayOpenBeforeDrag = true;

        if (this.shouldUseCompactTray()) {
            if (!this.pullTrayPeek) {
                this._trayWasExpandedBeforeDrag = true;
            }
            this.pullTrayPeek = true;
        } else {
            this._trayWasExpandedBeforeDrag = true;
            this.pullTrayOpen = false;
        }
    }

    beginDragFromPending() {
        const pending = this._pendingDrag;
        this.isPressingForDrag = false;
        this._pendingDrag = null;
        this.clearLongPressTimer();

        if (!pending) {
            return;
        }

        if (pending.type === 'event') {
            this.dragMode = 'event';
            this.draggingEventId = pending.id;
            this.draggingWorkOrderId = null;
            this.dragDayBodyTop = pending.dayBodyTop;
            this.dragStartDayIndex = pending.dayIndex;
            this.dragCurrentDayIndex = pending.dayIndex;
            this.dragStartLocal = pending.localStart;
            this.dragStartEndLocal = null;
            this.dragStartClientX = pending.clientX;
            this.dragStartClientY = pending.clientY;
            this.dragDayBodyHeight = pending.dayBodyHeight;
            this.dragDayWidth = pending.dayWidth;
            this.dragHasMoved = false;
            this.dragPreviewLocal = null;
            this.dragPreviewDurationHours = null;
            this.dragDurationHours = null;

            // Find the original event DOM node to clone its size
            const selector = `.sfs-calendar-day-body [data-id="${pending.id}"][data-day-index="${pending.dayIndex}"]`;
            const evtEl = this.template.querySelector(selector);
            let width = 120;
            let height = 40;

            if (evtEl) {
                const rect = evtEl.getBoundingClientRect();
                width = rect.width;
                height = rect.height;
            }

            // Work out the type class for coloring
            const appt = this.appointments.find(a => a.appointmentId === pending.id);
            const typeClass = appt
                ? this.getEventTypeClass(appt.workTypeName)
                : '';

            if (appt) {
                const endLocal = appt.schedEnd
                    ? this.convertUtcToUserLocal(appt.schedEnd)
                    : new Date(
                          pending.localStart.getTime() + 60 * 60 * 1000
                      );

                this.dragStartEndLocal = endLocal;

                const durationHours = this.computeDurationHours(
                    pending.localStart,
                    endLocal
                );
                this.dragDurationHours = durationHours;
                this.dragPreviewDurationHours = durationHours;

                const totalHours = this.calendarEndHour - this.calendarStartHour;
                const scaledHeight =
                    (durationHours / totalHours) * this.dragDayBodyHeight;
                height = scaledHeight;

                const timeLabel = this.formatTimeRange(
                    pending.localStart,
                    endLocal
                );

                this.showDragGhost(
                    pending.clientX,
                    pending.clientY,
                    pending.title,
                    timeLabel,
                    typeClass,
                    width,
                    height
                );

                this.compactTrayForDrag();

                this.updateSelectedEventStyles();
                this.showTrayCancelZone = false;
                this.registerGlobalDragListeners();
                return;
            }

            const timeLabel = pending.localStart.toLocaleTimeString([], {
                hour: 'numeric',
                minute: '2-digit'
            });

            this.showDragGhost(
                pending.clientX,
                pending.clientY,
                pending.title,
                timeLabel,
                typeClass,
                width,
                height
            );

            this.compactTrayForDrag();

            this.updateSelectedEventStyles();
        } else if (pending.type === 'wo') {
            this.dragMode = 'wo';
            this.draggingWorkOrderId = pending.workOrderId;
            this.draggingEventId = null;
            this.dragStartDayIndex = pending.dayIndex;
            this.dragCurrentDayIndex = pending.dayIndex;
            this.dragStartLocal = null;
            this.dragStartEndLocal = null;
            this.dragStartClientX = pending.clientX;
            this.dragStartClientY = pending.clientY;
            this.dragDayBodyHeight = pending.dayBodyHeight;
            this.dragDayWidth = pending.dayWidth;
            this.dragHasMoved = false;
            this.dragDayBodyTop = pending.dayBodyTop;
            this.dragPreviewLocal = null;
            this.dragPreviewDurationHours = this.defaultWorkOrderDurationHours;
            this.dragDurationHours = this.defaultWorkOrderDurationHours;


            // Approximate size for a new 6-hour event
            const dayHeight = pending.dayBodyHeight;
            const sixHourHeight =
                (this.defaultWorkOrderDurationHours / 24) * dayHeight;
            const width = pending.dayWidth * 0.88; // match left/right 6% padding
            const height = sixHourHeight;

            const dayObj = this.calendarDays[pending.dayIndex];
            const preview = new Date(dayObj.date);
            preview.setHours(9, 0, 0, 0);

            const endPreview = new Date(preview);
            endPreview.setHours(
                endPreview.getHours() + this.defaultWorkOrderDurationHours,
                endPreview.getMinutes(),
                0,
                0
            );

            const timeLabel = this.formatTimeRange(preview, endPreview);

            this.showDragGhost(
                pending.clientX,
                pending.clientY,
                pending.title,
                timeLabel,
                'sfs-event-default',
                width,
                height
            );
        }

        this.compactTrayForDrag();
        this.showTrayCancelZone = pending.type === 'wo';
        this.registerGlobalDragListeners();



    }

    showDragGhost(x, y, title, timeLabel, typeClass, width, height) {
        this.dragGhostVisible = true;
        this.dragGhostX = x;
        this.dragGhostY = y;
        this.dragGhostTitle = title || '';
        this.dragGhostTime = timeLabel || '';
        this.dragGhostTypeClass = typeClass || '';
        this.dragGhostWidth = width || 120;
        this.dragGhostHeight = height || 40;
    }

    hideDragGhost() {
        this.dragGhostVisible = false;
        this.dragGhostTitle = '';
        this.dragGhostTime = '';
        this.dragGhostTypeClass = '';
        this.dragGhostWidth = 0;
        this.dragGhostHeight = 0;
    }


    getDayIndexFromClientX(clientX) {
        const dayEls = Array.from(
            this.template.querySelectorAll('.sfs-calendar-day')
        );

        if (!dayEls.length) {
            return null;
        }

        let bestIndex = null;
        let bestDistance = Infinity;

        dayEls.forEach(el => {
            const idxStr = el.dataset.dayIndex;
            if (idxStr === undefined) {
                return;
            }

            const rect = el.getBoundingClientRect();
            const center = rect.left + rect.width / 2;

            if (clientX >= rect.left && clientX <= rect.right) {
                bestIndex = parseInt(idxStr, 10);
                bestDistance = 0;
                this.dragDayWidth = rect.width || this.dragDayWidth;
                return;
            }

            const distance = Math.abs(clientX - center);
            if (distance < bestDistance) {
                bestDistance = distance;
                bestIndex = parseInt(idxStr, 10);
                this.dragDayWidth = rect.width || this.dragDayWidth;
            }
        });

        return bestIndex;
    }


    stopAutoScrollLoop() {
        if (this._autoScrollFrame) {
            cancelAnimationFrame(this._autoScrollFrame);
            this._autoScrollFrame = null;
        }

        this._autoScrollPoint = null;
    }

    updateAutoScroll(clientX, clientY) {
        this._autoScrollPoint = { clientX, clientY };

        if (!this._autoScrollFrame) {
            this._autoScrollFrame = requestAnimationFrame(() =>
                this.performAutoScroll()
            );
        }
    }

    performAutoScroll() {
        this._autoScrollFrame = null;

        if (!this.dragMode || !this._autoScrollPoint) {
            return;
        }

        const { clientX, clientY } = this._autoScrollPoint;
        const edgeThreshold = 80;
        const maxStep = 18;
        const viewportWidth = window.innerWidth || 0;
        const viewportHeight = window.innerHeight || 0;

        const computeStep = distance => {
            const overlap = Math.max(edgeThreshold - distance, 0);
            if (!overlap) {
                return 0;
            }
            return Math.round((overlap / edgeThreshold) * maxStep);
        };

        let deltaX = 0;
        let deltaY = 0;

        const leftDistance = clientX;
        const rightDistance = viewportWidth - clientX;
        const topDistance = clientY;
        const bottomDistance = viewportHeight - clientY;

        deltaX -= computeStep(leftDistance);
        deltaX += computeStep(rightDistance);
        deltaY -= computeStep(topDistance);
        deltaY += computeStep(bottomDistance);

        const wrapper = this.template.querySelector('.sfs-calendar-days-wrapper');
        if (wrapper && deltaX !== 0) {
            wrapper.scrollLeft += deltaX;
        }

        if (deltaY !== 0) {
            window.scrollBy({ top: deltaY, behavior: 'auto' });
        }

        if (deltaX !== 0 || deltaY !== 0) {
            this._autoScrollFrame = requestAnimationFrame(() =>
                this.performAutoScroll()
            );
        }
    }




    handleCalendarPointerMove(event) {
        if (this.isCalendarPanMode) {
            return;
        }

        if (!this.dragMode || this.dragStartClientX === null) {
            return;
        }

        const clientPoint = this.getClientPoint(event);
        if (!clientPoint) {
            return;
        }
        const { clientX, clientY } = clientPoint;

        this.updateAutoScroll(clientX, clientY);
        this.updateTrayCancelHover(clientX, clientY);

        const dx = clientX - this.dragStartClientX;
        const dy = clientY - this.dragStartClientY;

        if (!this.dragHasMoved && (Math.abs(dx) > 5 || Math.abs(dy) > 5)) {
            this.dragHasMoved = true;
        }

        let newDayIndex = this.getDayIndexFromClientX(clientX);
        if (newDayIndex == null) {
            const dayOffset =
                this.dragMode === 'resize'
                    ? 0
                    : Math.round(dx / this.dragDayWidth);
            newDayIndex = this.dragStartDayIndex + dayOffset;
        }
        if (newDayIndex < 0) newDayIndex = 0;
        if (newDayIndex >= this.calendarDays.length) {
            newDayIndex = this.calendarDays.length - 1;
        }
        this.dragCurrentDayIndex = newDayIndex;

        const bodyHeightForDelta = this.dragDayBodyHeight || 1;
        let hoursDelta = (dy / bodyHeightForDelta) * 24;

        // When resizing, snap duration changes to 15-minute increments
        if (this.dragMode === 'resize') {
            hoursDelta = Math.round(hoursDelta / 0.25) * 0.25;
        }

        // Compute preview local time for the ghost
        let previewLocal;
        const dayObj = this.calendarDays[newDayIndex];
        const baseDay = new Date(dayObj.date);

        const dayEl = this.template.querySelector(
            `.sfs-calendar-day[data-day-index="${newDayIndex}"]`
        );
        const dayBodyEl = dayEl
            ? dayEl.querySelector('.sfs-calendar-day-body')
            : null;
        const bodyRect = dayBodyEl
            ? dayBodyEl.getBoundingClientRect()
            : null;

        if (bodyRect) {
            this.dragDayBodyHeight = bodyRect.height || this.dragDayBodyHeight;
            this.dragDayBodyTop = bodyRect.top;
        }

        if (this.dragMode === 'event' && this.dragStartLocal) {
            previewLocal = new Date(baseDay);
            previewLocal.setHours(
                this.dragStartLocal.getHours(),
                this.dragStartLocal.getMinutes(),
                0,
                0
            );

            const millisDelta = hoursDelta * 60 * 60 * 1000;
            previewLocal.setTime(previewLocal.getTime() + millisDelta);
        } else if (this.dragMode === 'resize' && this.dragStartLocal) {
            previewLocal = new Date(baseDay);
            previewLocal.setHours(
                this.dragStartLocal.getHours(),
                this.dragStartLocal.getMinutes(),
                0,
                0
            );
        } else if (this.dragMode === 'wo') {
            previewLocal = new Date(baseDay);
            const usableHeight = this.dragDayBodyHeight || 1;
            const relativeY = Math.max(
                0,
                Math.min(
                    usableHeight,
                    clientY - (this.dragDayBodyTop != null ? this.dragDayBodyTop : 0)
                )
            );
            const totalHours = this.calendarEndHour - this.calendarStartHour;
            const hourFraction =
                this.calendarStartHour + (relativeY / usableHeight) * totalHours;
            const hours = Math.floor(hourFraction);
            const minutes = Math.round((hourFraction - hours) * 60);
            previewLocal.setHours(hours, minutes, 0, 0);
        }

        if (previewLocal) {
            const minutes = previewLocal.getMinutes();
            const roundedMinutes = Math.round(minutes / 15) * 15;
            previewLocal.setMinutes(roundedMinutes, 0, 0);

            this.dragPreviewLocal = new Date(previewLocal);

        }

        const totalHours = this.calendarEndHour - this.calendarStartHour;

        if (previewLocal && this.dragDayBodyTop != null) {
            let durationHours = this.dragDurationHours || 1;

            if (this.dragMode === 'resize' && this.dragStartEndLocal) {
                const baseDuration = this.computeDurationHours(
                    this.dragStartLocal,
                    this.dragStartEndLocal
                );
                durationHours = baseDuration + hoursDelta;
                const startHourFraction =
                    previewLocal.getHours() + previewLocal.getMinutes() / 60;
                const maxDuration =
                    this.calendarEndHour - startHourFraction;
                durationHours = Math.min(
                    Math.max(durationHours, 0.25),
                    Math.max(maxDuration, 0.25)
                );
                this.dragPreviewDurationHours = durationHours;
            } else {
                this.dragPreviewDurationHours = durationHours;
            }

            const endPreview = new Date(previewLocal);
            endPreview.setTime(
                endPreview.getTime() + durationHours * 60 * 60 * 1000
            );

            const timeLabel = this.formatTimeRange(previewLocal, endPreview);

            // Convert previewLocal time -> vertical position in day body
            let hourFraction =
                previewLocal.getHours() + previewLocal.getMinutes() / 60;

            // Clamp within visible calendar hours
            if (hourFraction < this.calendarStartHour) {
                hourFraction = this.calendarStartHour;
            }
            if (hourFraction > this.calendarEndHour) {
                hourFraction = this.calendarEndHour;
            }

            const bodyHeight = this.dragDayBodyHeight || 1;
            const topRatio =
                (hourFraction - this.calendarStartHour) / totalHours;
            const yWithinBody = topRatio * bodyHeight;

            const ghostHeight = (durationHours / totalHours) * bodyHeight;

            // Position ghost horizontally at the center of the target day column
            let ghostX = clientX;
            const dayElForGhost = this.template.querySelector(
                `.sfs-calendar-day[data-day-index="${newDayIndex}"]`
            );
            if (dayElForGhost) {
                const dayRect = dayElForGhost.getBoundingClientRect();
                ghostX = dayRect.left + dayRect.width / 2;
                // Keep day width in sync
                this.dragDayWidth = dayRect.width;
            }

            // Top of the event box is the selected time; ghostY is the top
            const ghostY =
                (this.dragDayBodyTop != null ? this.dragDayBodyTop : 0) +
                yWithinBody; // top edge represents start time

            this.showDragGhost(
                ghostX,
                ghostY,
                this.dragGhostTitle,
                timeLabel,
                this.dragGhostTypeClass,
                this.dragGhostWidth,
                ghostHeight || this.dragGhostHeight
            );
        } else {
            // Fallback – just move with finger if we somehow lack geometry
            this.showDragGhost(
                clientX,
                clientY,
                this.dragGhostTitle,
                this.dragGhostTime,
                this.dragGhostTypeClass,
                this.dragGhostWidth,
                this.dragGhostHeight
            );
        }
        event.preventDefault();
    }


    handleCalendarPointerEnd(event) {
        if (this.isCalendarPanMode) {
            this.resetDragState();
            return;
        }

        if (!this.dragMode || this.dragStartClientX === null) {
            this.resetDragState();
            return;
        }

        const clientPoint = this.getClientPoint(event);
        if (!clientPoint) {
            this.resetDragState();
            return;
        }

        const { clientX, clientY } = clientPoint;

        const requiresExplicitPlacement = this.dragRequiresExplicitConfirmation;

        if (this.dragMode === 'wo' && this.isPointInTrayCancelZone(clientX, clientY)) {
            this.stopAutoScrollLoop();
            this.resetDragState();
            event.preventDefault();
            return;
        }

        const dy = clientY - this.dragStartClientY;
        const hoursDelta = (dy / this.dragDayBodyHeight) * 24;

        const finalDayIndex =
            this.dragCurrentDayIndex != null
                ? this.dragCurrentDayIndex
                : this.dragStartDayIndex;

        const dayObj = this.calendarDays[finalDayIndex];
        const baseDay = new Date(dayObj.date);

        // If user did not move enough, treat as no drag
        if (!this.dragHasMoved) {
            this.stopAutoScrollLoop();
            if (requiresExplicitPlacement && this.pendingSchedulePlacement) {
                this.freezeGhostForConfirmation();
            } else {
                this.resetDragState();
            }
            event.preventDefault();
            return;
        }

        if (this.dragMode === 'event') {
            const id = this.draggingEventId;
            const startLocal = this.dragStartLocal;

            const newLocal = new Date(baseDay);
            newLocal.setHours(
                startLocal.getHours(),
                startLocal.getMinutes(),
                0,
                0
            );

            const millisDelta = hoursDelta * 60 * 60 * 1000;
            newLocal.setTime(newLocal.getTime() + millisDelta);

            const minutes = newLocal.getMinutes();
            const roundedMinutes = Math.round(minutes / 15) * 15;
            newLocal.setMinutes(roundedMinutes, 0, 0);

            const isoString = this.toUserIsoString(newLocal);
            const duration =
                this.dragDurationHours ||
                this.dragPreviewDurationHours ||
                this.computeDurationHours(startLocal, this.dragStartEndLocal);
            const endLocal = new Date(newLocal);
            endLocal.setTime(endLocal.getTime() + duration * 60 * 60 * 1000);

            if (requiresExplicitPlacement) {
                this.cachePendingSchedulePlacement({
                    type: 'event',
                    appointmentId: id,
                    startIso: isoString,
                    endIso: this.toUserIsoString(endLocal),
                    dayIndex: finalDayIndex,
                    durationHours: duration,
                    title: this.dragGhostTitle,
                    typeClass: this.dragGhostTypeClass
                });
            } else {
                this.appointments = this.appointments.map(a => {
                    if (a.appointmentId === id) {
                        return {
                            ...a,
                            newStart: isoString,
                            disableSave: false
                        };
                    }
                    return a;
                });

                if (
                    this.selectedAppointment &&
                    this.selectedAppointment.appointmentId === id
                ) {
                    this.selectedAppointment = {
                        ...this.selectedAppointment,
                        newStart: isoString,
                        disableSave: false
                    };
                }

                this.handleReschedule({ target: { dataset: { id } } });
            }
        } else if (this.dragMode === 'resize') {
            const id = this.draggingEventId;
            if (id && this.dragStartLocal) {
                const startLocal = new Date(baseDay);
                startLocal.setHours(
                    this.dragStartLocal.getHours(),
                    this.dragStartLocal.getMinutes(),
                    0,
                    0
                );

                let durationHours =
                    this.dragPreviewDurationHours || this.dragDurationHours || 1;

                const newEnd = new Date(startLocal);
                newEnd.setTime(
                    newEnd.getTime() + durationHours * 60 * 60 * 1000
                );

                const roundedMinutes =
                    Math.round(newEnd.getMinutes() / 15) * 15;
                newEnd.setMinutes(roundedMinutes, 0, 0);

                this.updateAppointmentEndTime(id, this.toUserIsoString(newEnd));
            }
        } else if (this.dragMode === 'wo') {
            const workOrderId = this.draggingWorkOrderId;
            if (workOrderId) {
                const dayEl = this.template.querySelector(
                    `.sfs-calendar-day[data-day-index="${finalDayIndex}"]`
                );
                const dayBodyEl = dayEl
                    ? dayEl.querySelector('.sfs-calendar-day-body')
                    : null;
                const bodyRect = dayBodyEl
                    ? dayBodyEl.getBoundingClientRect()
                    : null;

                let dropLocal =
                    this.dragPreviewLocal != null
                        ? new Date(this.dragPreviewLocal)
                        : null;

                if (bodyRect) {
                    const usableHeight =
                        bodyRect.height || this.dragDayBodyHeight || 1;
                    const startHour = this.calendarStartHour;
                    const totalHours =
                        this.calendarEndHour - this.calendarStartHour;

                    const relativeY = Math.max(
                        0,
                        Math.min(
                            usableHeight,
                            clientY - (bodyRect.top || 0)
                        )
                    );

                    const hourFraction =
                        startHour + (relativeY / usableHeight) * totalHours;
                    const hours = Math.floor(hourFraction);
                    const minutes = Math.round((hourFraction - hours) * 60);

                    dropLocal = new Date(baseDay);
                    dropLocal.setHours(hours, minutes, 0, 0);
                }

                if (!dropLocal) {
                    dropLocal = new Date(baseDay);
                    dropLocal.setHours(9, 0, 0, 0);
                    const millisDelta = hoursDelta * 60 * 60 * 1000;
                    dropLocal.setTime(dropLocal.getTime() + millisDelta);
                }

                const minutes = dropLocal.getMinutes();
                const roundedMinutes = Math.round(minutes / 15) * 15;
                dropLocal.setMinutes(roundedMinutes, 0, 0);

                const dropEnd = new Date(dropLocal);
                const durationHours =
                    this.dragPreviewDurationHours ||
                    this.defaultWorkOrderDurationHours;
                dropEnd.setTime(
                    dropEnd.getTime() + durationHours * 60 * 60 * 1000
                );

                if (requiresExplicitPlacement) {
                    this.cachePendingSchedulePlacement({
                        type: 'wo',
                        workOrderId,
                        startIso: this.toUserIsoString(dropLocal),
                        endIso: this.toUserIsoString(dropEnd),
                        dayIndex: finalDayIndex,
                        durationHours,
                        title: this.dragGhostTitle,
                        typeClass: this.dragGhostTypeClass
                    });
                } else {
                    this.createAppointmentFromWorkOrder(
                        workOrderId,
                        this.toUserIsoString(dropLocal),
                        this.toUserIsoString(dropEnd)
                    );
                }
            }
        }

        if (requiresExplicitPlacement && this.showDragConfirmActions) {
            this.freezeGhostForConfirmation();
        } else {
            this.resetDragState();
        }
        event.preventDefault();
    }

    cachePendingSchedulePlacement(placement) {
        if (!placement) {
            return;
        }

        this.pendingSchedulePlacement = placement;
        this.isAwaitingScheduleConfirmation = true;
        this.updateGhostFromPlacement();
        this.attachGhostAnchorUpdater();
    }

    freezeGhostForConfirmation() {
        this.stopAutoScrollLoop();
        this.unregisterGlobalDragListeners();
        this.dragMode = null;
        this.dragStartClientX = null;
        this.dragStartClientY = null;
        this.dragHasMoved = false;
        this.showTrayCancelZone = false;
        this.isHoveringCancelZone = false;
        this.updateGhostFromPlacement();
    }

    confirmPendingSchedule() {
        const placement = this.pendingSchedulePlacement;
        if (!placement) {
            return;
        }

        this.pendingSchedulePlacement = null;
        this.isAwaitingScheduleConfirmation = false;
        this.dragRequiresExplicitConfirmation = false;

        if (placement.type === 'event') {
            const id = placement.appointmentId;
            if (!id || !placement.startIso) {
                this.resetDragState();
                return;
            }

            this.appointments = this.appointments.map(a => {
                if (a.appointmentId === id) {
                    return {
                        ...a,
                        newStart: placement.startIso,
                        disableSave: false
                    };
                }
                return a;
            });

            if (
                this.selectedAppointment &&
                this.selectedAppointment.appointmentId === id
            ) {
                this.selectedAppointment = {
                    ...this.selectedAppointment,
                    newStart: placement.startIso,
                    disableSave: false
                };
            }

            this.resetDragState();
            this.handleReschedule({ target: { dataset: { id } } });
            this.schedulePreviewCardId = null;
            this.schedulePreviewListMode = null;
            return;
        }

        if (placement.type === 'wo') {
            const { workOrderId, startIso, endIso } = placement;
            if (!workOrderId || !startIso || !endIso) {
                this.resetDragState();
                return;
            }

            this.resetDragState();
            this.createAppointmentFromWorkOrder(workOrderId, startIso, endIso);
            this.schedulePreviewCardId = null;
            this.schedulePreviewListMode = null;
        }
    }

    cancelPendingSchedule() {
        const cardId = this.schedulePreviewCardId;
        const listMode = this.schedulePreviewListMode;

        this.pendingSchedulePlacement = null;
        this.isAwaitingScheduleConfirmation = false;
        this.dragRequiresExplicitConfirmation = false;
        this.detachGhostAnchorUpdater();
        this.resetDragState();

        if (listMode) {
            this.listMode = listMode;
        }

        this.updateActiveTabState('list');

        if (cardId) {
            this.safeSetTimeout(() => this.scrollCardIntoView(cardId), 50);
        }

        this.schedulePreviewCardId = null;
        this.schedulePreviewListMode = null;
    }

    attachGhostAnchorUpdater() {
        if (this._boundGhostAnchorUpdater) {
            return;
        }

        this._boundGhostAnchorUpdater = () => this.updateGhostFromPlacement();

        const wrapper = this.template.querySelector('.sfs-calendar-days-wrapper');
        if (wrapper) {
            wrapper.addEventListener('scroll', this._boundGhostAnchorUpdater);
        }

        window.addEventListener('resize', this._boundGhostAnchorUpdater);
    }

    detachGhostAnchorUpdater() {
        if (!this._boundGhostAnchorUpdater) {
            return;
        }

        const wrapper = this.template.querySelector('.sfs-calendar-days-wrapper');
        if (wrapper) {
            wrapper.removeEventListener('scroll', this._boundGhostAnchorUpdater);
        }

        window.removeEventListener('resize', this._boundGhostAnchorUpdater);
        this._boundGhostAnchorUpdater = null;
    }

    updateGhostFromPlacement() {
        if (!this.pendingSchedulePlacement || !this.dragGhostVisible) {
            return;
        }

        const placement = this.pendingSchedulePlacement;
        const { startIso, endIso, dayIndex, durationHours, title, typeClass } =
            placement;

        const startLocal = startIso
            ? this.convertUtcToUserLocal(startIso)
            : null;
        const endLocal = endIso ? this.convertUtcToUserLocal(endIso) : null;

        const targetDayIndex =
            dayIndex != null ? dayIndex : this.resolveDayIndexFromDate(startLocal);

        if (targetDayIndex == null || !this.calendarDays[targetDayIndex]) {
            return;
        }

        const dayEl = this.template.querySelector(
            `.sfs-calendar-day[data-day-index="${targetDayIndex}"]`
        );
        const dayBodyEl = dayEl
            ? dayEl.querySelector('.sfs-calendar-day-body')
            : null;

        if (!dayEl || !dayBodyEl || !startLocal || !endLocal) {
            return;
        }

        const dayRect = dayEl.getBoundingClientRect();
        const bodyRect = dayBodyEl.getBoundingClientRect();
        const totalHours = this.calendarEndHour - this.calendarStartHour;
        const ghostDuration =
            durationHours || this.computeDurationHours(startLocal, endLocal);

        let hourFraction =
            startLocal.getHours() + startLocal.getMinutes() / 60;
        if (hourFraction < this.calendarStartHour) {
            hourFraction = this.calendarStartHour;
        }
        if (hourFraction > this.calendarEndHour) {
            hourFraction = this.calendarEndHour;
        }

        const bodyHeight = bodyRect.height || 1;
        const topRatio =
            (hourFraction - this.calendarStartHour) / totalHours;
        const yWithinBody = topRatio * bodyHeight;
        const ghostHeight = (ghostDuration / totalHours) * bodyHeight;

        const ghostX = dayRect.left + dayRect.width / 2;
        const ghostY = bodyRect.top + yWithinBody;

        this.showDragGhost(
            ghostX,
            ghostY,
            title || this.dragGhostTitle,
            this.formatTimeRange(startLocal, endLocal),
            typeClass || this.dragGhostTypeClass,
            this.dragGhostWidth,
            ghostHeight || this.dragGhostHeight
        );
    }

    resolveDayIndexFromDate(date) {
        if (!date || !this.calendarDays || !this.calendarDays.length) {
            return null;
        }

        const target = `${date.getFullYear()}-${this.pad2(
            date.getMonth() + 1
        )}-${this.pad2(date.getDate())}`;

        const matchIndex = this.calendarDays.findIndex(d => d.date === target);
        return matchIndex >= 0 ? matchIndex : null;
    }

    handleGhostPointerDown(event) {
        if (
            event.target &&
            event.target.closest('.sfs-drag-ghost__action')
        ) {
            return;
        }

        if (
            !this.isAwaitingScheduleConfirmation ||
            !this.pendingSchedulePlacement
        ) {
            return;
        }

        const point = this.getClientPoint(event);
        if (!point) {
            return;
        }

        event.preventDefault();
        event.stopPropagation();

        this.prepareGhostDragFromPlacement(point);
    }

    prepareGhostDragFromPlacement(startPoint) {
        const placement = this.pendingSchedulePlacement;
        if (!placement) {
            return;
        }

        const startLocal = placement.startIso
            ? this.convertUtcToUserLocal(placement.startIso)
            : null;
        const endLocal = placement.endIso
            ? this.convertUtcToUserLocal(placement.endIso)
            : null;

        const durationHours = placement.durationHours
            ? placement.durationHours
            : this.computeDurationHours(startLocal, endLocal);

        const dayIndex =
            placement.dayIndex != null
                ? placement.dayIndex
                : this.resolveDayIndexFromDate(startLocal);

        if (dayIndex == null) {
            return;
        }

        this.dragMode = placement.type === 'event' ? 'event' : 'wo';
        this.dragRequiresExplicitConfirmation = true;
        this.dragStartDayIndex = dayIndex;
        this.dragCurrentDayIndex = dayIndex;
        this.dragStartLocal = startLocal;
        this.dragStartEndLocal = endLocal;
        this.dragPreviewLocal = startLocal;
        this.dragPreviewDurationHours = durationHours;
        this.dragDurationHours = durationHours;
        this.dragStartClientX = startPoint.clientX;
        this.dragStartClientY = startPoint.clientY;
        this.dragHasMoved = false;
        this.isAwaitingScheduleConfirmation = true;

        if (placement.type === 'event') {
            this.draggingEventId = placement.appointmentId;
            this.draggingWorkOrderId = null;
        } else {
            this.draggingWorkOrderId = placement.workOrderId;
            this.draggingEventId = null;
        }

        const dayEl = this.template.querySelector(
            `.sfs-calendar-day[data-day-index="${dayIndex}"]`
        );
        const bodyEl = dayEl
            ? dayEl.querySelector('.sfs-calendar-day-body')
            : null;
        const dayRect = dayEl ? dayEl.getBoundingClientRect() : null;
        const bodyRect = bodyEl ? bodyEl.getBoundingClientRect() : null;

        if (dayRect && bodyRect) {
            this.dragDayBodyHeight = bodyRect.height || this.dragDayBodyHeight;
            this.dragDayBodyTop = bodyRect.top;
            this.dragDayWidth = dayRect.width || this.dragDayWidth;
        }

        this.registerGlobalDragListeners();
    }

    scrollCardIntoView(cardId) {
        if (!cardId) {
            return;
        }

        const cardSelector = `.sfs-card[data-card-id="${cardId}"]`;
        const cardEl = this.template.querySelector(cardSelector);

        if (cardEl && typeof cardEl.scrollIntoView === 'function') {
            cardEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }

    resetDragState() {
        this.dragMode = null;
        this.draggingEventId = null;
        this.draggingWorkOrderId = null;
        this.dragStartClientX = null;
        this.dragStartClientY = null;
        this.dragStartDayIndex = null;
        this.dragCurrentDayIndex = null;
        this.dragDayWidth = null;
        this.dragDayBodyHeight = null;
        this.dragStartLocal = null;
        this.dragStartEndLocal = null;
        this.dragPreviewLocal = null;
        this.dragPreviewDurationHours = null;
        this.dragDurationHours = null;
        this.dragHasMoved = false;
        this.isPressingForDrag = false;
        this._pendingDrag = null;
        this.clearLongPressTimer();
        this.hideDragGhost();
        this.dragDayBodyTop = null;
        this.showTrayCancelZone = false;
        this.isHoveringCancelZone = false;
        this.dragRequiresExplicitConfirmation = false;
        this.isAwaitingScheduleConfirmation = false;
        this.pendingSchedulePlacement = null;
        this.stopAutoScrollLoop();
        this.unregisterGlobalDragListeners();
        this.detachGhostAnchorUpdater();
        this.updateSelectedEventStyles();

        if (this._trayOpenBeforeDrag) {
            this.pullTrayOpen = true;
        }

        if (this.pullTrayOpen && this._trayWasExpandedBeforeDrag) {
            this.pullTrayPeek = false;
        }

        this._trayWasExpandedBeforeDrag = false;
        this._trayOpenBeforeDrag = false;
    }



    // ======= DATA LOAD =======

    loadAppointments() {
        this.isLoading = true;
        this.selectedAppointment = null;
        this.selectedAbsence = null;

        getMyAppointmentsOnline({ targetUserId: this.activeUserId })
            .then(result => {
                const appts =
                    result && result.appointments ? result.appointments : [];

                this.userTimeZoneId =
                    result && result.userTimeZoneId ? result.userTimeZoneId : null;
                this.userTimeZoneShort = this.computeTimeZoneShort();

                this.currentUserId =
                    result && result.debug ? result.debug.currentUserId : null;
                this.activeUserId =
                    result && result.viewingUserId
                        ? result.viewingUserId
                        : this.activeUserId || this.currentUserId;
                this.managerTeam =
                    result && result.teamMembers ? result.teamMembers : [];
                this.isManager = Boolean(result && result.isManager);
                this.viewingUserName = this.resolveViewingName();

                if (this.isViewingAsOther) {
                    const selected = this.managerTeam.find(
                        m => m.userId === this.activeUserId
                    );
                    this.selectedManagerUserId = selected
                        ? selected.userId
                        : this.selectedManagerUserId;
                    this.selectedManagerUserName = selected
                        ? selected.name
                        : this.selectedManagerUserName;
                } else {
                    this.selectedManagerUserId = null;
                    this.selectedManagerUserName = '';
                }

                this.debugInfo =
                    result && result.debug
                        ? result.debug
                        : { note: 'No debug info.' };

                // Unscheduled Work Orders for tray
                const unscheduled =
                    result && result.unscheduledWorkOrders
                        ? result.unscheduledWorkOrders
                        : [];
                this.unscheduledWorkOrders = unscheduled.map(wo => {
                    const clone = { ...wo };
                    clone.isCrewAppointment = Boolean(wo.hasCrewAssignment);
                    clone.serviceAppointmentCount = wo.serviceAppointmentCount || 0;
                    clone.unscheduledServiceAppointmentCount =
                        wo.unscheduledServiceAppointmentCount || 0;
                    clone.allServiceAppointmentsReturnRequired = Boolean(
                        wo.allServiceAppointmentsReturnRequired
                    );
                    clone.hasUnscheduledNonReturnAppointment = Boolean(
                        wo.hasUnscheduledNonReturnAppointment
                    );
                    clone.hasAnyReturnVisitRequired = Boolean(
                        wo.hasAnyReturnVisitRequired
                    );
                    clone.recordTypeName = wo.recordTypeName;
                    clone.needsReturnVisitScheduling = Boolean(
                        wo.needsReturnVisitScheduling
                    );
                    clone.workOrderSubject = wo.subject;
                    clone.cardId = `wo-${wo.workOrderId}`;
                    clone.hasAppointment = false;
                    clone.workTypeName = wo.workTypeName;
                    clone.contactPhoneHref = wo.contactPhone
                        ? `tel:${wo.contactPhone.replace(/\D/g, '')}`
                        : null;
                    clone.contactEmailHref = wo.contactEmail
                        ? `mailto:${encodeURIComponent(wo.contactEmail)}`
                        : null;

                    const reporter = this.parseReporterInfo(
                        wo.reporterContactInfo
                    );
                    clone.reporterContactInfo = wo.reporterContactInfo;
                    clone.reporterName = reporter.name;
                    clone.reporterPhone = reporter.phone;
                    clone.reporterPhoneDisplay = reporter.phoneDisplay;
                    clone.reporterPhoneHref = reporter.phoneHref;
                    clone.reporterEmail = reporter.email;
                    clone.reporterEmailHref = reporter.emailHref;

                    const parts = [
                        wo.street,
                        wo.city,
                        wo.state,
                        wo.postalCode,
                        wo.country
                    ].filter(Boolean);
                    clone.fullAddress = parts.join(', ');
                    return clone;
                });

                const transferRequests =
                    result && result.transferRequests
                        ? result.transferRequests
                        : [];
                const submittedTransferRequests =
                    result && result.submittedTransferRequests
                        ? result.submittedTransferRequests
                        : [];

                this.transferRequests = transferRequests.map(req =>
                    this.normalizeTransferRequest(req)
                );
                this.submittedTransferRequests = submittedTransferRequests
                    .filter(req => !req.acceptedOn && !req.rejectedOn)
                    .map(req => this.normalizeTransferRequest(req));

                const absences =
                    result && result.absences ? result.absences : [];
                this.absences = absences.map(abs => ({
                    ...abs,
                    subject: abs.description || 'Absence',
                    newStart: abs.start,
                    newEnd: abs.endTime
                }));

                this.appointments = appts.map(appt => {
                    const clone = { ...appt };

                    clone.newStart = appt.schedStart;
                    clone.disableSave = true;
                    clone.workOrderId = appt.workOrderId;
                    clone.recordUrl =
                        '/lightning/r/ServiceAppointment/' +
                        appt.appointmentId +
                        '/view';

                    clone.workOrderStatus = appt.workOrderStatus;

                    clone.isExpanded = false;

                    const typeClass = this.getEventTypeClass(appt.workTypeName);
                    clone.workTypeClass = `sfs-worktype ${typeClass || ''}`.trim();

                    if (appt.contactPhone) {
                        const digits = appt.contactPhone.replace(/\D/g, '');
                        clone.contactPhoneHref = digits ? `tel:${digits}` : null;
                    } else {
                        clone.contactPhoneHref = null;
                    }

                    if (appt.contactEmail) {
                        clone.contactEmailHref = `mailto:${encodeURIComponent(
                            appt.contactEmail
                        )}`;
                    } else {
                        clone.contactEmailHref = null;
                    }

                    clone.reporterContactInfo = appt.reporterContactInfo;

                    const reporter = this.parseReporterInfo(
                        appt.reporterContactInfo
                    );
                    clone.reporterName = reporter.name;
                    clone.reporterPhone = reporter.phone;
                    clone.reporterPhoneDisplay = reporter.phoneDisplay;
                    clone.reporterPhoneHref = reporter.phoneHref;
                    clone.reporterEmail = reporter.email;
                    clone.reporterEmailHref = reporter.emailHref;
                    clone.workOrderNumber = appt.workOrderNumber;


                    const parts = [
                        appt.street,
                        appt.city,
                        appt.state,
                        appt.postalCode,
                        appt.country
                    ].filter(Boolean);
                    clone.fullAddress = parts.join(', ');

                    const crewMembers = appt.crewMembers || [];
                    clone.crewMembers = crewMembers;
                    clone.crewOptions = crewMembers.map(m => ({
                        label: m.name,
                        value: m.serviceResourceId
                    }));
                    clone.selectedCrewMemberId = null;
                    clone.disableAssignTech = true;

                    clone.quoteAttachmentUrl = appt.quoteAttachmentDownloadUrl || null;
                    clone.quoteAttachmentDocumentId =
                        appt.quoteAttachmentDocumentId || null;
                    clone.hasQuoteAttachment = Boolean(
                        appt.hasQuoteAttachment ||
                            appt.workOrderStatus === 'Quote Attached'
                    );
                    clone.showQuoteActions =
                        clone.workOrderStatus === 'Quote Sent' ||
                        this.isQuoteAttachedAppointment(clone);
                    clone.showMarkQuoteSentAction = this.isQuoteAttachedAppointment(
                        clone
                    );
                    clone.opportunityRecordType = appt.opportunityRecordType || null;
                    clone.cardId = appt.appointmentId;
                    clone.hasAppointment = true;

                    return clone;
                });

                if (!this.timelineStartDate && !this.weekStartDate) {
                    this.centerCalendarOnToday();
                } else {
                    this.buildCalendarModel();
                }
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling getMyAppointmentsOnline',
                    errorMessage: message
                };
                this.appointments = [];
                this.calendarDays = [];
                this.unscheduledWorkOrders = [];
                this.absences = [];
                this.showToast('Error loading appointments', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    computeTimeZoneShort() {
        if (!this.userTimeZoneId) {
            return null;
        }
        try {
            const formatter = new Intl.DateTimeFormat('en-US', {
                timeZone: this.userTimeZoneId,
                timeZoneName: 'short'
            });
            const parts = formatter.formatToParts(new Date());
            const tzPart = parts.find(p => p.type === 'timeZoneName');
            return tzPart ? tzPart.value : this.userTimeZoneId;
        } catch (e) {
            return this.userTimeZoneId;
        }
    }

    resolveViewingName() {
        if (!this.isManager || !this.activeUserId) {
            return 'You';
        }

        if (this.currentUserId && this.activeUserId === this.currentUserId) {
            return 'You';
        }

        const match = (this.managerTeam || []).find(
            m => m.userId === this.activeUserId
        );
        return match ? match.name : 'Team Member';
    }

    convertUtcToUserLocal(dateLike) {
        const d = new Date(dateLike);

        if (!this.userTimeZoneId) {
            return d;
        }

        try {
            const formatter = new Intl.DateTimeFormat('en-US', {
                timeZone: this.userTimeZoneId,
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false
            });

            const parts = formatter.formatToParts(d).reduce((acc, part) => {
                acc[part.type] = part.value;
                return acc;
            }, {});

            return new Date(
                Number(parts.year),
                Number(parts.month) - 1,
                Number(parts.day),
                Number(parts.hour),
                Number(parts.minute),
                Number(parts.second),
                d.getMilliseconds()
            );
        } catch (e) {
            return d;
        }
    }

    getTimeZoneOffsetMinutes(date) {
        try {
            const formatter = new Intl.DateTimeFormat('en-US', {
                timeZone: this.userTimeZoneId,
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false
            });

            const parts = formatter.formatToParts(date).reduce((acc, part) => {
                acc[part.type] = part.value;
                return acc;
            }, {});

            const tzAsUtc = Date.UTC(
                Number(parts.year),
                Number(parts.month) - 1,
                Number(parts.day),
                Number(parts.hour),
                Number(parts.minute),
                Number(parts.second)
            );

            return (tzAsUtc - date.getTime()) / (60 * 1000);
        } catch (e) {
            // Default to the browser's current offset when we cannot parse the timezone
            return -date.getTimezoneOffset();
        }
    }

    toUserIsoString(dateLike) {
        const local = new Date(dateLike);

        if (!this.userTimeZoneId) {
            return local.toISOString();
        }

        // Adjust only when the browser's timezone differs from the user's
        const userOffsetMinutes = this.getTimeZoneOffsetMinutes(local);
        const browserOffsetMinutes = -local.getTimezoneOffset();
        const offsetDeltaMinutes = userOffsetMinutes - browserOffsetMinutes;

        if (offsetDeltaMinutes === 0) {
            return local.toISOString();
        }

        const utcMs = local.getTime() - offsetDeltaMinutes * 60 * 1000;
        return new Date(utcMs).toISOString();
    }

    getUserNow() {
        if (!this.userTimeZoneId) {
            return new Date();
        }

        try {
            const formatter = new Intl.DateTimeFormat('en-US', {
                timeZone: this.userTimeZoneId,
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false
            });

            const parts = formatter.formatToParts(new Date()).reduce(
                (acc, part) => {
                    acc[part.type] = part.value;
                    return acc;
                },
                {}
            );

            return new Date(
                Number(parts.year),
                Number(parts.month) - 1,
                Number(parts.day),
                Number(parts.hour),
                Number(parts.minute),
                Number(parts.second)
            );
        } catch (e) {
            return new Date();
        }
    }

    computeDurationHours(startDate, endDate) {
        if (!startDate || !endDate) {
            return 1;
        }
        const diffMs = endDate.getTime() - startDate.getTime();
        return Math.max(diffMs / (60 * 60 * 1000), 0.25);
    }

    formatTimeRange(startDate, endDate) {
        if (!startDate || !endDate) {
            return '';
        }

        const startText = startDate.toLocaleTimeString([], {
            hour: 'numeric',
            minute: '2-digit'
        });
        const endText = endDate.toLocaleTimeString([], {
            hour: 'numeric',
            minute: '2-digit'
        });

        return `${startText} – ${endText}`;
    }

    pad2(num) {
        return num.toString().padStart(2, '0');
    }

    formatPhoneDigits(digits) {
        if (!digits) {
            return null;
        }
        const onlyDigits = digits.replace(/\D/g, '');
        if (!onlyDigits) {
            return null;
        }

        if (onlyDigits.length === 10) {
            return (
                onlyDigits.slice(0, 3) +
                ' ' +
                onlyDigits.slice(3, 6) +
                ' ' +
                onlyDigits.slice(6)
            );
        }
        if (onlyDigits.length === 11 && onlyDigits.startsWith('1')) {
            return (
                '+1 ' +
                onlyDigits.slice(1, 4) +
                ' ' +
                onlyDigits.slice(4, 7) +
                ' ' +
                onlyDigits.slice(7)
            );
        }
        return onlyDigits;
    }

    parseReporterInfo(raw) {
        const result = {
            name: null,
            phone: null,
            phoneDisplay: null,
            phoneHref: null,
            email: null,
            emailHref: null
        };

        if (!raw) {
            return result;
        }

        const text = String(raw).trim();
        if (!text) {
            return result;
        }

        const emailMatch = text.match(
            /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i
        );
        if (emailMatch) {
            result.email = emailMatch[0];
            result.emailHref = `mailto:${encodeURIComponent(result.email)}`;
        }

        const digits = text.replace(/\D/g, '');
        if (digits && digits.length >= 7) {
            result.phone = digits;
            result.phoneDisplay = this.formatPhoneDigits(digits);
            result.phoneHref = `tel:${digits}`;
        }

        let nameSource = text;
        if (result.email) {
            nameSource = nameSource.replace(result.email, ' ');
        }

        nameSource = nameSource.replace(/[\d\-\+\(\)\.]/g, ' ');

        const name = nameSource.replace(/\s+/g, ' ').trim();
        result.name = name || null;

        return result;
    }

    buildCalendarModel() {
        if (!this.timelineStartDate && !this.weekStartDate) {
            this.centerCalendarOnToday();
            return;
        }

        this.showNowLine = false;
        this.nowLineStyle = '';

        const days = [];
        const dayMap = new Map();
        this.todayDayIndex = null;

        const startBase = this.isTimelineMode
            ? new Date(this.timelineStartDate || new Date())
            : new Date(this.weekStartDate || new Date());

        startBase.setHours(0, 0, 0, 0);
        const daysToShow = this.daysToShow;

        const totalHours = this.calendarEndHour - this.calendarStartHour;

        const nowLocal = this.getUserNow();

        for (let i = 0; i < daysToShow; i++) {
            const d = new Date(startBase);
            d.setDate(startBase.getDate() + i);

            const year = d.getFullYear();
            const month = d.getMonth() + 1;
            const dayNum = d.getDate();

            const key =
                year + '-' + this.pad2(month) + '-' + this.pad2(dayNum);

            const isToday =
                year === nowLocal.getFullYear() &&
                month === nowLocal.getMonth() + 1 &&
                dayNum === nowLocal.getDate();

            const weekdayLabel = d.toLocaleDateString([], {
                weekday: 'short'
            });
            const fullLabel = d.toLocaleDateString([], {
                weekday: 'short',
                month: 'short',
                day: 'numeric'
            });

            const day = {
                key,
                date: d,
                events: [],
                isToday,
                weekdayLabel,
                dayNumberLabel: dayNum,
                fullLabel,
                cssClassTimeline: isToday
                    ? 'sfs-calendar-day sfs-calendar-day_today'
                    : 'sfs-calendar-day',
                cssClassWeekCell: isToday
                    ? 'sfs-week-cell sfs-week-cell_today'
                    : 'sfs-week-cell'
            };

            if (isToday) {
                this.todayDayIndex = i;
            }

            days.push(day);
            dayMap.set(key, day);
        }

        this.appointments.forEach(appt => {
            if (!appt.schedStart) {
                return;
            }

            const startLocal = this.convertUtcToUserLocal(appt.schedStart);
            const endLocal = appt.schedEnd
                ? this.convertUtcToUserLocal(appt.schedEnd)
                : this.convertUtcToUserLocal(appt.schedStart);

            const year = startLocal.getFullYear();
            const month = startLocal.getMonth() + 1;
            const dayNum = startLocal.getDate();

            const dayKey =
                year + '-' + this.pad2(month) + '-' + this.pad2(dayNum);

            const day = dayMap.get(dayKey);
            if (!day) {
                return;
            }

            const eventKey = `${appt.appointmentId}-${dayKey}`;

            let startHour =
                startLocal.getHours() + startLocal.getMinutes() / 60;
            let endHour =
                endLocal.getHours() + endLocal.getMinutes() / 60;

            startHour = Math.max(startHour, this.calendarStartHour);
            endHour = Math.min(endHour, this.calendarEndHour);

            if (endHour <= startHour) {
                endHour = startHour + 0.25;
            }

            const topPct =
                ((startHour - this.calendarStartHour) / totalHours) * 100;
            const heightPct =
                ((endHour - startHour) / totalHours) * 100;

            const timeLabel = startLocal.toLocaleTimeString([], {
                hour: 'numeric',
                minute: '2-digit'
            });

            const displaySubject = appt.workOrderNumber
                ? `${appt.workOrderNumber} — ${appt.workOrderSubject}`
                : appt.workOrderSubject;

            const typeClass = this.getEventTypeClass
                ? this.getEventTypeClass(appt.workTypeName)
                : '';
            const baseTimelineClass = `sfs-calendar-event ${typeClass}`.trim();
            const baseWeekClass = `sfs-week-event-box ${typeClass}`.trim();

            day.events.push({
                id: appt.appointmentId,
                key: eventKey,
                style: `top:${topPct}%;height:${heightPct}%;`,
                subject: appt.workOrderSubject,
                workOrderNumber: appt.workOrderNumber,
                workTypeName: appt.workTypeName,
                timeLabel,
                isCrewAssignment: appt.isCrewAssignment,
                isMyAssignment: appt.isMyAssignment,
                isAbsence: false,
                kind: 'appointment',
                className: baseTimelineClass,
                classNameWeek: baseWeekClass,
                baseClassTimeline: baseTimelineClass,
                baseClassWeek: baseWeekClass
            });
        });

        this.absences.forEach(abs => {
            if (!abs.start || !abs.endTime) {
                return;
            }

            const startLocal = this.convertUtcToUserLocal(abs.start);
            const endLocal = this.convertUtcToUserLocal(abs.endTime);

            const startDay = new Date(startLocal);
            startDay.setHours(0, 0, 0, 0);
            const endDay = new Date(endLocal);
            endDay.setHours(0, 0, 0, 0);

            const baseTimelineClass = 'sfs-calendar-event sfs-calendar-event_absence';
            const baseWeekClass = 'sfs-week-event-box sfs-week-event-box_absence';

            for (
                let cursor = new Date(startDay);
                cursor.getTime() <= endDay.getTime();
                cursor.setDate(cursor.getDate() + 1)
            ) {
                const year = cursor.getFullYear();
                const month = cursor.getMonth() + 1;
                const dayNum = cursor.getDate();
                const dayKey =
                    year + '-' + this.pad2(month) + '-' + this.pad2(dayNum);
                const day = dayMap.get(dayKey);
                if (!day) {
                    continue;
                }

                const dayStart = new Date(cursor);
                const dayEnd = new Date(cursor);
                dayEnd.setHours(23, 59, 59, 999);

                const segmentStart =
                    startLocal.getTime() > dayStart.getTime()
                        ? startLocal
                        : dayStart;
                const segmentEnd =
                    endLocal.getTime() < dayEnd.getTime() ? endLocal : dayEnd;

                let startHour =
                    segmentStart.getHours() + segmentStart.getMinutes() / 60;
                let endHour = segmentEnd.getHours() + segmentEnd.getMinutes() / 60;

                startHour = Math.max(startHour, this.calendarStartHour);
                endHour = Math.min(endHour, this.calendarEndHour);

                if (endHour <= startHour) {
                    endHour = startHour + 0.25;
                }

                const topPct =
                    ((startHour - this.calendarStartHour) / totalHours) * 100;
                const heightPct = ((endHour - startHour) / totalHours) * 100;

                const timeLabel = this.formatTimeRange(segmentStart, segmentEnd);
                const key = `${abs.absenceId}-${dayKey}`;

                day.events.push({
                    id: abs.absenceId,
                    key,
                    style: `top:${topPct}%;height:${heightPct}%;`,
                    subject: abs.subject || 'Absence',
                    workOrderNumber: null,
                    workTypeName: null,
                    timeLabel,
                    isCrewAssignment: false,
                    isMyAssignment: false,
                    isAbsence: true,
                    kind: 'absence',
                    className: baseTimelineClass,
                    classNameWeek: baseWeekClass,
                    baseClassTimeline: baseTimelineClass,
                    baseClassWeek: baseWeekClass
                });
            }
        });

        if (this.isTimelineMode) {
            this.showNowLine = true;
        }

        this.calendarDays = days;
        this.updateSelectedEventStyles();
        this.scheduleNowLinePositionUpdate();
    }

    scheduleNowLinePositionUpdate() {
        if (
            !this.hasWindow ||
            typeof window.requestAnimationFrame !== 'function' ||
            typeof window.cancelAnimationFrame !== 'function'
        ) {
            return;
        }

        if (this._nowLineFrame) {
            window.cancelAnimationFrame(this._nowLineFrame);
        }

        this._nowLineFrame = window.requestAnimationFrame(() => {
            this._nowLineFrame = null;
            this.updateNowLinePosition();
        });
    }

    updateNowLinePosition() {
        if (!this.isTimelineMode) {
            this.showNowLine = false;
            this.nowLineStyle = '';
            return;
        }

        const calendarEl = this.template.querySelector('.sfs-calendar');
        const dayBodyEl = this.template.querySelector('.sfs-calendar-day-body');

        if (!calendarEl || !dayBodyEl) {
            return;
        }

        const calendarRect = calendarEl.getBoundingClientRect();
        const bodyRect = dayBodyEl.getBoundingClientRect();

        const totalHours = this.calendarEndHour - this.calendarStartHour;
        const nowLocal = this.getUserNow();
        const nowHourFraction =
            nowLocal.getHours() + nowLocal.getMinutes() / 60;

        const clampedHour = Math.min(
            Math.max(nowHourFraction, this.calendarStartHour),
            this.calendarEndHour
        );

        const relativePct =
            (clampedHour - this.calendarStartHour) / totalHours;

        const offsetTop = bodyRect.top - calendarRect.top;
        const topPx = offsetTop + relativePct * bodyRect.height;

        const style = `top:${topPx}px;`;

        if (!this.showNowLine || this.nowLineStyle !== style) {
            this.showNowLine = true;
            this.nowLineStyle = style;
        }
    }

    updateSelectedEventStyles() {
        const selectedIds = [];
        if (this.selectedAppointment) {
            selectedIds.push(this.selectedAppointment.appointmentId);
        }

        if (this.selectedAbsence) {
            selectedIds.push(this.selectedAbsence.absenceId);
        }

        const draggingId = this.dragMode ? this.draggingEventId : null;

        const updatedDays = this.calendarDays.map(day => {
            const newEvents = day.events.map(evt => {
                const baseTimeline =
                    evt.baseClassTimeline || 'sfs-calendar-event';
                const baseWeek =
                    evt.baseClassWeek || 'sfs-week-event-box';

                const isSelected =
                    selectedIds.length > 0 && selectedIds.includes(evt.id);
                const isDragging = draggingId && evt.id === draggingId;

                let classTimeline = baseTimeline;
                let classWeek = baseWeek;

                if (isSelected) {
                    classTimeline += ' sfs-calendar-event_selected';
                    classWeek += ' sfs-week-event-box_selected';
                }

                if (isDragging) {
                    classTimeline += ' sfs-calendar-event_dragging';
                    classWeek += ' sfs-week-event-box_dragging';
                }

                return {
                    ...evt,
                    className: classTimeline,
                    classNameWeek: classWeek
                };
            });

            return { ...day, events: newEvents };
        });

        this.calendarDays = updatedDays;
    }


    // ======= LIST TAB HANDLERS =======

    handleListModeChange(event) {
        const mode = event?.detail?.value || event.target?.dataset?.mode;
        this.setListMode(mode);
    }

    setListMode(mode) {
        if (!mode || mode === this.listMode) {
            return;
        }
        this.listMode = mode;
    }

    handleQuoteAttachmentClick(event) {
        event.preventDefault();

        const cardId = event.currentTarget.dataset.id;
        if (!cardId) {
            return;
        }

        const appt = this.findAppointmentByCardId(cardId);

        if (!appt) {
            return;
        }

        const docId =
            appt.quoteAttachmentDocumentId ||
            this.extractContentDocumentId(appt.quoteAttachmentUrl);
        const downloadUrl =
            this.normalizeDownloadUrl(appt.quoteAttachmentUrl) ||
            this.buildDownloadUrlFromDocId(docId);

        if (!docId && !downloadUrl) {
            this.showToast(
                'Quote unavailable',
                'We could not find a quote attachment for this appointment.',
                'warning'
            );
            return;
        }

        if (docId) {
            // Use the native file previewer so the experience matches opening the
            // attachment from the Files related list in the mobile app. If
            // navigation fails (e.g. invalid record id for preview), fall back
            // to the raw download URL so the tech can still access the file.
            try {
                const navPromise = this[NavigationMixin.Navigate]({
                    type: 'standard__namedPage',
                    attributes: {
                        pageName: 'filePreview'
                    },
                    state: {
                        recordIds: docId,
                        selectedRecordId: docId
                    }
                });

                if (navPromise && typeof navPromise.catch === 'function') {
                    navPromise.catch(error =>
                        this.handleNavigationError(downloadUrl, error)
                    );
                }
            } catch (err) {
                this.handleNavigationError(downloadUrl, err);
            }

            return;
        }

        this.navigateToDownload(downloadUrl);
    }

    handleNavigationError(fallbackUrl, error) {
        // If we have a direct download URL, use it as a fallback so the tech
        // can still view the quote file. Otherwise, surface a toast so the user
        // knows navigation failed instead of seeing a generic routing error.
        if (fallbackUrl) {
            this.navigateToDownload(fallbackUrl);
            return;
        }

        const message =
            (error &&
                (error.message ||
                    (error.body && error.body.message))) ||
            'Unable to open the quote attachment.';
        this.showToast('Navigation failed', message, 'error');
    }

    navigateToDownload(targetUrl) {
        if (!targetUrl) {
            this.showToast(
                'Attachment unavailable',
                'Unable to locate the quote attachment.',
                'warning'
            );
            return;
        }

        this[NavigationMixin.Navigate]({
            type: 'standard__webPage',
            attributes: { url: targetUrl }
        });
    }

    normalizeDownloadUrl(url) {
        if (!url) {
            return null;
        }

        // Mobile apps require absolute URLs; the server returns a relative path
        // (e.g. "/sfc/servlet.shepherd/document/download/<docId>"). Prefix with
        // the current origin so the in-app browser can resolve it.
        if (url.startsWith('/') && this.hasWindow) {
            return `${window.location.origin}${url}`;
        }

        return url;
    }

    buildDownloadUrlFromDocId(docId) {
        if (!docId) {
            return null;
        }

        if (!this.hasWindow) {
            return null;
        }

        return `${window.location.origin}/sfc/servlet.shepherd/document/download/${docId}`;
    }

    extractContentDocumentId(url) {
        if (!url) {
            return null;
        }

        // Extract the 15 or 18 character ContentDocumentId from a shepherd
        // download or view URL so we can open the native file preview.
        const match = url.match(/\/document\/(?:download|view)\/([a-zA-Z0-9]{15,18})/);
        return match ? match[1] : null;
    }

    handleMarkQuoteSent(event) {
        const workOrderId = event.target.dataset.woid;
        if (!workOrderId) {
            return;
        }

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to update the work order status.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        markWorkOrderQuoteSent({ workOrderId })
            .then(() => {
                this.showToast(
                    'Status updated',
                    'Work order marked as Quote Sent.',
                    'success'
                );
                return this.loadAppointments();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling markWorkOrderQuoteSent',
                    errorMessage: message
                };
                this.showToast('Error updating status', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleMyStatusChange(event) {
        this.selectedMyStatus = event.detail.value;
    }

    handleCrewMemberChange(event) {
        const id = event.target.dataset.id;
        const value = event.detail.value;

        this.appointments = this.appointments.map(appt => {
            if (appt.appointmentId === id) {
                return {
                    ...appt,
                    selectedCrewMemberId: value,
                    disableAssignTech: !value
                };
            }
            return appt;
        });

        if (
            this.selectedAppointment &&
            this.selectedAppointment.appointmentId === id
        ) {
            this.selectedAppointment = {
                ...this.selectedAppointment,
                selectedCrewMemberId: value,
                disableAssignTech: !value
            };
        }
    }

    handleAssignToCrewMember(event) {
        const id = event.target.dataset.id;
        const appt = this.appointments.find(a => a.appointmentId === id);

        if (!appt || !appt.selectedCrewMemberId) {
            this.showToast(
                'Pick a technician',
                'Select a crew member before assigning.',
                'warning'
            );
            return;
        }

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to assign an appointment.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        assignCrewAppointment({
            serviceAppointmentId: id,
            serviceResourceId: appt.selectedCrewMemberId
        })
            .then(() => {
                this.showToast(
                    'Appointment reassigned',
                    'The appointment has been assigned to the selected technician.',
                    'success'
                );
                this.loadAppointments();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling assignCrewAppointment',
                    errorMessage: message
                };
                this.showToast('Error assigning appointment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleToggleListDetails(event) {
        const id = event.currentTarget.dataset.id;
        this.appointments = this.appointments.map(appt => {
            if (appt.appointmentId === id) {
                return {
                    ...appt,
                    isExpanded: !appt.isExpanded
                };
            }
            return appt;
        });
    }

    handleRefresh() {
        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to refresh appointments.',
                'warning'
            );
            return;
        }
        this.loadAppointments();
    }

    handleDateChange(event) {
        const id = event.target.dataset.id;
        const value = event.target.value;

        this.appointments = this.appointments.map(appt => {
            if (appt.appointmentId === id) {
                const changed = value && value !== appt.schedStart;
                return {
                    ...appt,
                    newStart: value,
                    disableSave: !changed
                };
            }
            return appt;
        });

        if (
            this.selectedAppointment &&
            this.selectedAppointment.appointmentId === id
        ) {
            const changed = value && value !== this.selectedAppointment.schedStart;
            this.selectedAppointment = {
                ...this.selectedAppointment,
                newStart: value,
                disableSave: !changed
            };
        }
    }

    handleReschedule(event) {
        const id = event.target.dataset.id;
        const appt = this.appointments.find(a => a.appointmentId === id);

        if (!appt || !appt.newStart) {
            this.showToast(
                'Pick a date and time',
                'Select a new start date and time before rescheduling.',
                'warning'
            );
            return;
        }

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to reschedule an appointment.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        rescheduleAppointment({ appointmentId: id, newStart: appt.newStart })
            .then(() => {
                this.showToast(
                    'Appointment updated',
                    'The appointment has been rescheduled.',
                    'success'
                );
                return this.loadAppointments();
            })
            .then(() => {
                this.handleCalendarToday();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling rescheduleAppointment',
                    errorMessage: message
                };
                this.showToast('Error updating appointment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleScheduleOnCalendar(event) {
        const cardId =
            event && event.currentTarget && event.currentTarget.dataset
                ? event.currentTarget.dataset.id
                : null;

        if (!cardId) {
            return;
        }

        const target = this.findAppointmentByCardId(cardId);

        if (!target) {
            return;
        }

        this.schedulePreviewCardId = target.cardId || null;
        this.schedulePreviewListMode = this.listMode;
        this.isAwaitingScheduleConfirmation = false;
        this.pendingSchedulePlacement = null;
        this.dragRequiresExplicitConfirmation = true;

        this.startScheduleOnCalendar(target);
    }

    startScheduleOnCalendar(target) {
        if (!target) {
            return;
        }

        this.updateActiveTabState('calendar');
        this.ensureCalendarReadyForScheduling(target);
    }

    ensureCalendarReadyForScheduling(target, attempt = 0) {
        const maxAttempts = 6;
        const dayEl = this.template.querySelector('.sfs-calendar-day');
        const dayBodyEl = this.template.querySelector('.sfs-calendar-day-body');

        if (!dayEl || !dayBodyEl) {
            if (attempt >= maxAttempts) {
                return;
            }

            this.safeSetTimeout(
                () => this.ensureCalendarReadyForScheduling(target, attempt + 1),
                120
            );

            return;
        }

        this.beginSchedulingGhost(target);
    }

    beginSchedulingGhost(target) {
        if (!target) {
            return;
        }

        this.isAwaitingScheduleConfirmation = false;
        this.pendingSchedulePlacement = null;

        const shouldRequireExplicitConfirmation =
            this.dragRequiresExplicitConfirmation;

        const dayIndex = this.resolveCalendarDayIndex(
            target.hasAppointment ? target.schedStart : null
        );

        const dayEl =
            this.template.querySelector(
                `.sfs-calendar-day[data-day-index="${dayIndex}"]`
            ) || this.template.querySelector('.sfs-calendar-day');

        const dayBodyEl = dayEl
            ? dayEl.querySelector('.sfs-calendar-day-body')
            : this.template.querySelector('.sfs-calendar-day-body');

        if (!dayEl || !dayBodyEl) {
            return;
        }

        const bodyRect = dayBodyEl.getBoundingClientRect();
        const clientX = bodyRect.left + bodyRect.width / 2;
        const clientY = bodyRect.top + bodyRect.height * 0.25;
        const pending = target.hasAppointment
            ? this.buildPendingEventFromList(
                  target,
                  dayIndex,
                  dayEl.clientWidth || bodyRect.width || 1,
                  bodyRect.height || dayBodyEl.clientHeight || 1,
                  bodyRect.top,
                  clientX,
                  clientY
              )
            : this.buildPendingWorkOrderFromList(
                  target,
                  dayIndex,
                  dayEl.clientWidth || bodyRect.width || 1,
                  bodyRect.height || dayBodyEl.clientHeight || 1,
                  bodyRect.top,
                  clientX,
                  clientY
              );

        if (!pending) {
            return;
        }

        this.resetDragState();
        this.dragRequiresExplicitConfirmation = shouldRequireExplicitConfirmation;
        this._pendingDrag = pending;
        this.beginDragFromPending();
    }

    buildPendingEventFromList(
        appt,
        dayIndex,
        dayWidth,
        dayBodyHeight,
        dayBodyTop,
        clientX,
        clientY
    ) {
        if (!appt || !appt.appointmentId) {
            return null;
        }

        const localStart = appt.schedStart
            ? this.convertUtcToUserLocal(appt.schedStart)
            : new Date();

        return {
            type: 'event',
            id: appt.appointmentId,
            dayIndex,
            localStart,
            clientX,
            clientY,
            dayBodyHeight,
            dayBodyTop,
            dayWidth,
            title: appt.workOrderSubject || appt.subject || 'Appointment'
        };
    }

    buildPendingWorkOrderFromList(
        workOrder,
        dayIndex,
        dayWidth,
        dayBodyHeight,
        dayBodyTop,
        clientX,
        clientY
    ) {
        if (!workOrder || !workOrder.workOrderId) {
            return null;
        }

        const title = workOrder.workOrderNumber
            ? `${workOrder.workOrderNumber} — ${workOrder.workOrderSubject ||
                  workOrder.subject ||
                  'New appointment'}`
            : workOrder.workOrderSubject || workOrder.subject || 'New appointment';

        return {
            type: 'wo',
            workOrderId: workOrder.workOrderId,
            dayIndex,
            clientX,
            clientY,
            dayBodyHeight,
            dayBodyTop,
            dayWidth,
            title
        };
    }

    resolveCalendarDayIndex(startDateLike) {
        if (this.calendarDays && this.calendarDays.length && startDateLike) {
            const targetDay = this.normalizeDayStart(
                this.convertUtcToUserLocal(startDateLike)
            );

            const index = this.calendarDays.findIndex(day => {
                const dayDate = this.normalizeDayStart(new Date(day.date));
                return dayDate && targetDay && dayDate.getTime() === targetDay.getTime();
            });

            if (index >= 0) {
                return index;
            }
        }

        if (this.todayDayIndex != null) {
            return this.todayDayIndex;
        }

        return 0;
    }

    normalizeDayStart(date) {
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
            return null;
        }

        const clone = new Date(date);
        clone.setHours(0, 0, 0, 0);
        return clone;
    }

    updateAppointmentEndTime(appointmentId, newEndIso) {
        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to change the appointment duration.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        updateAppointmentEnd({
            appointmentId,
            newEnd: newEndIso
        })
            .then(() => {
                this.showToast(
                    'Appointment updated',
                    'The appointment duration has been changed.',
                    'success'
                );
                return this.loadAppointments();
            })
            .then(() => {
                this.handleCalendarToday();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error updating appointment end time',
                    errorMessage: message
                };
                this.showToast('Error updating appointment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // ======= DETAIL RESCHEDULE HANDLERS =======

    handleDetailDateChange(event) {
        const value = event.target.value;
        if (!this.selectedAppointment) {
            return;
        }

        const id = this.selectedAppointment.appointmentId;
        const changed = value && value !== this.selectedAppointment.schedStart;

        this.selectedAppointment = {
            ...this.selectedAppointment,
            newStart: value,
            disableSave: !changed
        };

        this.appointments = this.appointments.map(appt => {
            if (appt.appointmentId === id) {
                return {
                    ...appt,
                    newStart: value,
                    disableSave: !changed
                };
            }
            return appt;
        });
    }

    handleDetailReschedule() {
        if (!this.selectedAppointment || !this.selectedAppointment.newStart) {
            this.showToast(
                'Pick a date and time',
                'Select a new start date and time before rescheduling.',
                'warning'
            );
            return;
        }

        const id = this.selectedAppointment.appointmentId;

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to reschedule an appointment.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        rescheduleAppointment({
            appointmentId: id,
            newStart: this.selectedAppointment.newStart
        })
            .then(() => {
                this.showToast(
                    'Appointment updated',
                    'The appointment has been rescheduled.',
                    'success'
                );
                return this.loadAppointments();
            })
            .then(() => {
                this.handleCalendarToday();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling rescheduleAppointment',
                    errorMessage: message
                };
                this.showToast('Error updating appointment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleListSubjectClick(event) {
        this.handleEventClick(event);
    }

    // ======= EVENT CLICK =======

    handleEventClick(event) {
        if (this.dragHasMoved) {
            this.dragHasMoved = false;
            return;
        }

        const id = event.currentTarget.dataset.id;
        const kind = event.currentTarget.dataset.kind || 'appointment';
        if (!id) {
            return;
        }

        if (kind === 'absence') {
            const absence = this.findAbsenceById(id);
            this.safeClearTimeout(this._closeTimeout);
            this.isDetailClosing = false;
            this.selectedAppointment = null;
            this.selectedAbsence = absence
                ? { ...absence, newStart: absence.newStart || absence.start, newEnd: absence.newEnd || absence.endTime }
                : null;
            this.updateSelectedEventStyles();
            return;
        }

        const appt = this.findAppointmentByCardId(id);

        this.safeClearTimeout(this._closeTimeout);
        this.isDetailClosing = false;
        this.selectedAbsence = null;
        this.selectedAppointment = appt ? { ...appt } : null;
        this.updateSelectedEventStyles();
    }

    // ======= DETAIL OVERLAY =======

    handleCloseDetails(event) {
        if (event) {
            event.stopPropagation();
        }

        if (!this.selectedAppointment || this.isDetailClosing) {
            return;
        }

        this.isDetailClosing = true;

        this.safeClearTimeout(this._closeTimeout);
        this._closeTimeout = this.safeSetTimeout(() => {
            this.selectedAppointment = null;
            this.isDetailClosing = false;
            this.updateSelectedEventStyles();
        }, 200);
    }

    handleDetailCardClick(event) {
        event.stopPropagation();
    }


    // ======= ABSENCE EDITING =======

    handleAbsenceDateChange(event) {
        if (!this.selectedAbsence) {
            return;
        }

        const field = event.target.dataset.field;
        if (!field) {
            return;
        }

        const value = event.detail ? event.detail.value : event.target.value;
        this.selectedAbsence = { ...this.selectedAbsence, [field]: value };
    }

    handleSaveAbsence() {
        if (!this.selectedAbsence) {
            return;
        }

        const { absenceId, newStart, newEnd } = this.selectedAbsence;
        this.isLoading = true;

        updateResourceAbsence({
            absenceId,
            startDateTimeIso: newStart,
            endDateTimeIso: newEnd
        })
            .then(() => {
                this.showToast(
                    'Absence updated',
                    'Absence time updated.',
                    'success'
                );
                this.selectedAbsence = null;
                return this.loadAppointments();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Error updating absence', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleDeleteAbsenceClick(event) {
        if (event) {
            event.stopPropagation();
        }

        const id = event && event.currentTarget ? event.currentTarget.dataset.id : null;
        const absenceId = id || (this.selectedAbsence ? this.selectedAbsence.absenceId : null);

        if (!absenceId) {
            return;
        }

        this.isLoading = true;

        deleteResourceAbsence({ absenceId })
            .then(() => {
                this.showToast(
                    'Absence deleted',
                    'The absence has been removed from the calendar.',
                    'success'
                );
                if (
                    this.selectedAbsence &&
                    this.selectedAbsence.absenceId === absenceId
                ) {
                    this.selectedAbsence = null;
                }
                return this.loadAppointments();
            })
            .then(() => {
                this.handleCalendarToday();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Error deleting absence', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleCloseAbsenceDetails(event) {
        if (event) {
            event.stopPropagation();
        }

        if (!this.selectedAbsence || this.isAbsenceDetailClosing) {
            return;
        }

        this.isAbsenceDetailClosing = true;

        this.safeClearTimeout(this._absenceCloseTimeout);
        this._absenceCloseTimeout = this.safeSetTimeout(() => {
            this.selectedAbsence = null;
            this.isAbsenceDetailClosing = false;
            this.updateSelectedEventStyles();
        }, 200);
    }

    navigateToRecord(recordId, objectApiName) {
        if (!recordId) {
            return;
        }

        if (this.isDesktopFormFactor) {
            this[NavigationMixin.Navigate]({
                type: 'standard__recordPage',
                attributes: {
                    recordId,
                    objectApiName,
                    actionName: 'view'
                }
            });
        } else {
            const deepLink = `com.salesforce.fieldservice://v1/sObject/${recordId}`;

            this[NavigationMixin.Navigate]({
                type: 'standard__webPage',
                attributes: {
                    url: deepLink
                }
            });
        }
    }

    handleOpenWorkOrder() {
        if (
            !this.selectedAppointment ||
            !this.selectedAppointment.appointmentId
        ) {
            return;
        }

        this.navigateToRecord(
            this.selectedAppointment.appointmentId,
            'ServiceAppointment'
        );
    }

    handleOpenAccount() {
        if (
            !this.selectedAppointment ||
            !this.selectedAppointment.accountId
        ) {
            return;
        }

        this.navigateToRecord(this.selectedAppointment.accountId, 'Account');
    }

    handleOpenContact() {
        if (
            !this.selectedAppointment ||
            !this.selectedAppointment.contactId
        ) {
            return;
        }

        this.navigateToRecord(this.selectedAppointment.contactId, 'Contact');
    }

    // ======= CALENDAR TAB HANDLERS =======

    handleManagerUserChange(event) {
        this.selectedManagerUserId = event.detail.value || null;
        const selected = (this.managerTeam || []).find(
            m => m.userId === this.selectedManagerUserId
        );
        this.selectedManagerUserName = selected ? selected.name : '';
    }

    handleManagerApply() {
        const targetId = this.selectedManagerUserId || this.currentUserId;
        this.activeUserId = targetId;
        this.viewingUserName = this.resolveViewingName();
        this.loadAppointments();
    }

    handleManagerReset() {
        this.selectedManagerUserId = null;
        this.selectedManagerUserName = '';
        this.activeUserId = this.currentUserId;
        this.viewingUserName = this.resolveViewingName();
        this.loadAppointments();
    }

    handleTabButtonClick(event) {
        const tabValue = event.currentTarget.dataset.tab;
        if (!tabValue) {
            return;
        }

        if (tabValue === 'manager' && !this.isManager) {
            this.updateActiveTabState('list');
            return;
        }

        this.updateActiveTabState(tabValue);
    }

    handleCalendarTabClick() {
        this.updateActiveTabState('calendar');
    }

    handleCalendarTabKeydown(event) {
        const isActivationKey = event.key === 'Enter' || event.key === ' ';
        if (!isActivationKey) {
            return;
        }

        event.preventDefault();
        this.handleCalendarTabClick();
    }

    handleCalendarPrev() {
        const step = this.isTimelineMode ? -1 : -7;
        this.shiftCalendar(step);
    }

    handleCalendarNext() {
        const step = this.isTimelineMode ? 1 : 7;
        this.shiftCalendar(step);
    }

    handleCalendarToday() {
        this.centerCalendarOnToday();

        if (this.isTimelineMode) {
            this._needsCenterOnToday = true;

            this.safeClearTimeout(this._centerTimeout);
            this._centerTimeout = this.safeSetTimeout(() => {
                this.centerTimelineOnTodayColumn();
            }, 0);
        }
    }

    handleCalendarMode(event) {
        const mode = event.target.dataset.mode;
        if (!mode) return;
        this.calendarMode = mode;

        if (!this.timelineStartDate && !this.weekStartDate) {
            this.centerCalendarOnToday();
        } else {
            this.buildCalendarModel();
        }

        if (this.calendarMode === 'timeline') {
            this._needsCenterOnToday = true;
        }
    }

    handleCalendarPanToggle() {
        this.isCalendarPanMode = !this.isCalendarPanMode;

        if (this.isCalendarPanMode) {
            this.resetDragState();
            this.isPressingForDrag = false;
            this._pendingDrag = null;
            this.clearLongPressTimer();
        }
    }

    // ======= TRAY HANDLERS =======

    toggleTray() {
        const willOpen = !this.pullTrayOpen;
        this.pullTrayOpen = willOpen;

        if (willOpen) {
            this.pullTrayPeek = this.shouldUseCompactTray();
        }
    }

    handleTrayPeekToggle() {
        this.pullTrayPeek = !this.pullTrayPeek;
        this._trayWasExpandedBeforeDrag = false;
    }

    shouldUseCompactTray() {
        if (typeof window !== 'undefined' && window.matchMedia) {
            return window.matchMedia('(max-width: 768px)').matches;
        }

        return !this.isDesktopFormFactor;
    }

    handleTrayInfoPointer(event) {
        if (event && typeof event.stopPropagation === 'function') {
            event.stopPropagation();
        }
        this.clearLongPressTimer();
        this.isPressingForDrag = false;
        this._pendingDrag = null;
    }

    handleTrayCardClick(event) {
        // Cancel any pending long-press drag so a quick tap does not start a drag
        this.clearLongPressTimer();
        this.isPressingForDrag = false;
        this._pendingDrag = null;

        // If we already transitioned into drag mode, ignore the click
        if (this.dragMode === 'wo' || this.dragHasMoved) {
            this.dragHasMoved = false;
            return;
        }

        const workOrderId =
            event && event.currentTarget && event.currentTarget.dataset
                ? event.currentTarget.dataset.woid
                : null;

        if (!workOrderId) {
            return;
        }

        this.navigateToWorkOrderInformation(workOrderId);
    }

    handleTrayCardInfoClick(event) {
        this.handleTrayInfoPointer(event);

        const workOrderId =
            event && event.currentTarget && event.currentTarget.dataset
                ? event.currentTarget.dataset.woid
                : null;

        if (!workOrderId) {
            return;
        }

        const workOrder = (this.unscheduledWorkOrders || []).find(
            wo => wo.workOrderId === workOrderId
        );

        if (!workOrder) {
            return;
        }

        const detail = this.normalizeWorkOrderDetail(workOrder);

        if (!detail) {
            return;
        }

        this.safeClearTimeout(this._closeTimeout);
        this.isDetailClosing = false;
        this.selectedAbsence = null;
        this.selectedAppointment = detail;
        this.updateSelectedEventStyles();
    }

    navigateToWorkOrderInformation(workOrderId) {
        if (!workOrderId) {
            return;
        }

        try {
            if (this.isDesktopFormFactor) {
                // Hint the lightning record page to show the Information tab.
                this[NavigationMixin.Navigate]({
                    type: 'standard__recordPage',
                    attributes: {
                        recordId: workOrderId,
                        objectApiName: 'WorkOrder',
                        actionName: 'view'
                    },
                    state: {
                        // Some orgs label the primary tab "Information"; when present, this
                        // deep-link lands the user there while still working for custom layouts.
                        tabsetName: 'Information'
                    }
                });
                return;
            }

            // In the mobile app, deep-link directly to the Information tab for consistency with
            // the dispatcher experience shown in the work order tray.
            const infoUrl =
                'com.salesforce.fieldservice://v1/sObject/' +
                workOrderId +
                '/information';
            infoUrl = `com.salesforce.fieldservice://v1/sObject/${workOrderId}/information`;

            // Navigate to the info URL
            this[NavigationMixin.Navigate]({
                type: 'standard__webPage',
                attributes: {
                    url: infoUrl
                }
            });

            try {
                // Navigate to the Work Order record page
                this[NavigationMixin.Navigate]({
                    type: 'standard__recordPage',
                    attributes: {
                        recordId: workOrderId,
                        objectApiName: 'WorkOrder',
                        actionName: 'view'
                    }
                });
            } catch (err) {
                const message =
                    (err &&
                        (err.message ||
                            (err.body && err.body.message))) ||
                    'Unable to open the work order details.';

                this.showToast('Navigation failed', message, 'error');
            }
        }
        catch (err) {
                const message =
                    (err &&
                        (err.message ||
                            (err.body && err.body.message))) ||
                    'Unable to open the work order details.';

                this.showToast('Navigation failed', message, 'error');
            }
    }

    createAppointmentFromWorkOrder(workOrderId, isoStart, isoEnd) {
        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to schedule an appointment.',
                'warning'
            );
            return;
        }

        this.isLoading = true;

        createAppointmentForWorkOrder({
            workOrderId,
            startDateTimeIso: isoStart,
            endDateTimeIso: isoEnd,
            targetUserId: this.activeUserId
        })
            .then(() => {
                this.showToast(
                    'Appointment created',
                    'A new appointment has been scheduled from this work order.',
                    'success'
                );
                return this.loadAppointments();
            })
            .then(() => {
                this.handleCalendarToday();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling createAppointmentForWorkOrder',
                    errorMessage: message
                };
                this.showToast('Error creating appointment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleUnassignClick(event) {
        event.stopPropagation();
        const id = event.currentTarget.dataset.id;
        if (!id) {
            return;
        }

        const targetAppt = this.appointments.find(
            appt => appt.appointmentId === id
        );

        this.unassignTarget = targetAppt
            ? {
                  id,
                  subject: targetAppt.subject,
                  workOrderNumber: targetAppt.workOrderNumber
              }
            : { id };

        this.isUnassignModalOpen = true;
    }

    closeUnassignModal() {
        this.isUnassignModalOpen = false;
        this.unassignTarget = null;
    }

    get unassignModalSubject() {
        if (!this.unassignTarget) {
            return '';
        }

        return this.unassignTarget.subject || 'Service Appointment';
    }

    get unassignModalWorkOrder() {
        if (!this.unassignTarget || !this.unassignTarget.workOrderNumber) {
            return '';
        }

        return `WO # ${this.unassignTarget.workOrderNumber}`;
    }

    confirmUnassignAppointment() {
        if (!this.unassignTarget || !this.unassignTarget.id) {
            return;
        }

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to update an assignment.',
                'warning'
            );
            return;
        }

        this.isLoading = true;
        const appointmentId = this.unassignTarget.id;

        unassignAppointment({ appointmentId })
            .then(() => {
                this.showToast(
                    'Removed from schedule',
                    'The appointment was returned to the scheduling queue.',
                    'success'
                );
                return this.loadAppointments();
            })
            .then(() => {
                if (this.isCalendarTabActive) {
                    this.handleCalendarToday();
                }
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.debugInfo = {
                    note: 'Error calling unassignAppointment',
                    errorMessage: message
                };
                this.showToast('Error removing assignment', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
                this.closeUnassignModal();
            });
    }

    handleRequestReschedule(event) {
        let workOrderId = event.currentTarget.dataset.woid;

        // Fallback to appointment lookup when only the appointment id is present
        if (!workOrderId && event.currentTarget.dataset.id) {
            const appt = this.appointments.find(
                a => a.appointmentId === event.currentTarget.dataset.id
            );
            workOrderId = appt ? appt.workOrderId : null;
        }

        if (!workOrderId) {
            return;
        }
        event.stopPropagation();
        this.openRescheduleModal(workOrderId);
    }

    openRescheduleModal(workOrderId) {
        this.isRescheduleModalOpen = true;
        this.rescheduleLoading = true;
        this.rescheduleOptions = [];
        this.rescheduleSelection = null;
        this.rescheduleWorkOrderId = workOrderId;

        getTerritoryResources({ workOrderId })
            .then(options => {
                const optionList = (options || [])
                    .filter(opt => opt.userId)
                    .map(opt => {
                        return {
                            label: opt.name,
                            value: opt.userId
                        };
                    });

                this.rescheduleOptions = optionList;

                if (this.rescheduleOptions.length) {
                    this.rescheduleSelection = this.rescheduleOptions[0].value;
                } else {
                    this.showToast(
                        'No technicians available',
                        'No active resources with linked users were found for this work order territory.',
                        'warning'
                    );
                }
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Unable to load resources', message, 'error');
            })
            .finally(() => {
                this.rescheduleLoading = false;
            });
    }

    closeRescheduleModal() {
        this.isRescheduleModalOpen = false;
        this.rescheduleWorkOrderId = null;
        this.rescheduleOptions = [];
        this.rescheduleSelection = null;
        this.rescheduleLoading = false;
    }

    handleRescheduleSelection(event) {
        this.rescheduleSelection = event.detail.value;
    }

    submitRescheduleRequest() {
        if (!this.rescheduleWorkOrderId || !this.rescheduleSelection) {
            return;
        }

        this.rescheduleLoading = true;

        createEngineerTransferRequest({
            workOrderId: this.rescheduleWorkOrderId,
            targetUserId: this.rescheduleSelection
        })
            .then(() => {
                this.showToast(
                    'Transfer request sent',
                    'The selected technician will review the transfer request.',
                    'success'
                );
                this.closeRescheduleModal();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Error requesting transfer', message, 'error');
            })
            .finally(() => {
                this.rescheduleLoading = false;
            });
    }

    handleAcceptTransferRequest(event) {
        const requestId = event.currentTarget.dataset.id;
        if (!requestId) {
            return;
        }

        this.isLoading = true;

        acceptEngineerTransferRequest({
            transferRequestId: requestId,
            targetOwnerId: this.activeUserId
        })
            .then(() => {
                this.showToast(
                    'Transfer accepted',
                    'The work order has been moved to your scheduling queue.',
                    'success'
                );
                return this.loadAppointments();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Unable to accept transfer', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    openRejectModal(event) {
        const requestId = event.currentTarget.dataset.id;
        if (!requestId) {
            return;
        }

        this.rejectRequestId = requestId;
        this.rejectReason = '';
        this.isRejectModalOpen = true;
    }

    closeRejectModal() {
        this.isRejectModalOpen = false;
        this.rejectReason = '';
        this.rejectRequestId = null;
    }

    handleRejectReasonChange(event) {
        this.rejectReason = event.target.value;
    }

    get rejectSubmitDisabled() {
        return this.isLoading || !this.rejectReason;
    }

    submitRejectRequest() {
        if (!this.rejectRequestId || !this.rejectReason) {
            return;
        }

        this.isLoading = true;

        rejectEngineerTransferRequest({
            transferRequestId: this.rejectRequestId,
            reason: this.rejectReason
        })
            .then(() => {
                this.showToast(
                    'Transfer rejected',
                    'The requester will be notified of the rejection reason.',
                    'success'
                );
                this.transferRequests = (this.transferRequests || []).filter(
                    req => req.transferRequestId !== this.rejectRequestId
                );
                this.closeRejectModal();
                return this.loadAppointments();
            })
            .catch(error => {
                const message = this.reduceError(error);
                this.showToast('Unable to reject transfer', message, 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    normalizeTransferRequest(req) {
        const clone = { ...req };
        const parts = [
            req.street,
            req.city,
            req.state,
            req.postalCode,
            req.country
        ].filter(Boolean);

        clone.fullAddress = parts.join(', ');

        const statusMeta = this.getTransferStatusMeta(req);
        clone.statusLabel = statusMeta.statusLabel;
        clone.statusClass = statusMeta.statusClass;
        clone.rejectionReason = req.rejectionReason;

        return clone;
    }

    getTransferStatusMeta(req) {
        if (req.acceptedOn) {
            return {
                statusLabel: 'Accepted',
                statusClass: 'sfs-status sfs-status_success'
            };
        }

        if (req.rejectedOn) {
            return {
                statusLabel: 'Rejected',
                statusClass: 'sfs-status sfs-status_error'
            };
        }

        return {
            statusLabel: 'Pending',
            statusClass: 'sfs-status sfs-status_pending'
        };
    }

    getEventTypeClass(workTypeName) {
        if (!workTypeName || typeof workTypeName !== 'string') {
            return 'sfs-event-default';
        }
        const name = workTypeName.toLowerCase();

        if (name.includes('break') || name.includes('fix')) {
            return 'sfs-event-breakfix';
        }
        if (
            name.includes('pm') ||
            name.includes('preventive') ||
            name.includes('preventative')
        ) {
            return 'sfs-event-pm';
        }
        if (name.includes('install')) {
            return 'sfs-event-install';
        }
        return 'sfs-event-default';
    }

    // ======= UTIL =======

    registerGlobalErrorHandlers() {
        if (
            this._hasRegisteredErrorHandlers ||
            !this.hasWindow ||
            typeof window.addEventListener !== 'function'
        ) {
            return;
        }

        this._boundOnGlobalError = event => this.onGlobalError(event);
        this._boundOnUnhandledRejection = event =>
            this.onUnhandledRejection(event);

        window.addEventListener('error', this._boundOnGlobalError);
        window.addEventListener(
            'unhandledrejection',
            this._boundOnUnhandledRejection
        );

        this._hasRegisteredErrorHandlers = true;
    }

    unregisterGlobalErrorHandlers() {
        if (!this._hasRegisteredErrorHandlers || !this.hasWindow) {
            return;
        }

        if (
            typeof window.removeEventListener === 'function' &&
            this._boundOnGlobalError
        ) {
            window.removeEventListener('error', this._boundOnGlobalError);
        }

        if (
            typeof window.removeEventListener === 'function' &&
            this._boundOnUnhandledRejection
        ) {
            window.removeEventListener(
                'unhandledrejection',
                this._boundOnUnhandledRejection
            );
        }

        this._hasRegisteredErrorHandlers = false;
        this._boundOnGlobalError = null;
        this._boundOnUnhandledRejection = null;
    }

    onGlobalError(event) {
        if (!event) {
            return;
        }

        if (typeof event.preventDefault === 'function') {
            event.preventDefault();
        }

        const error =
            event.error ||
            new Error(
                event.message || 'Script error occurred before error object.'
            );

        this.captureError(error, 'window.onerror');
    }

    onUnhandledRejection(event) {
        if (!event) {
            return;
        }

        if (typeof event.preventDefault === 'function') {
            event.preventDefault();
        }

        const reason = event.reason || event.detail?.reason;
        const error =
            reason instanceof Error
                ? reason
                : new Error(
                      reason || 'Unhandled promise rejection with no reason'
                  );

        this.captureError(error, 'window.unhandledrejection');
    }

    get hasWindow() {
        return typeof window !== 'undefined';
    }

    safeSetTimeout(callback, delay) {
        if (!this.hasWindow || typeof window.setTimeout !== 'function') {
            return null;
        }
        return window.setTimeout(callback, delay);
    }

    safeClearTimeout(handle) {
        if (this.hasWindow && typeof window.clearTimeout === 'function') {
            window.clearTimeout(handle);
        }
    }

    captureError(error, context = '') {
        if (typeof console !== 'undefined' && typeof console.error === 'function') {
            console.error('[fslHello]', context, error);
        }

        const message =
            (error && (error.message || error.body?.message)) || 'Unknown error';

        this.debugInfo = {
            ...this.debugInfo,
            lastError: {
                context,
                message,
                stack: error?.stack || null
            }
        };
    }

    reduceError(error) {
        let message = 'Unknown error';
        if (error && Array.isArray(error.body)) {
            message = error.body.map(e => e.message).join(', ');
        } else if (error && error.body && error.body.message) {
            message = error.body.message;
        } else if (error && error.message) {
            message = error.message;
        }
        return message;
    }

    showToast(title, message, variant) {
        this.dispatchEvent(
            new ShowToastEvent({
                title,
                message,
                variant
            })
        );
    }
}
