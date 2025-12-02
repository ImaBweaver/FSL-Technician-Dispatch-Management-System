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
import { ShowToastEvent } from 'lightning/platformShowToastEvent';

export default class FslHello extends NavigationMixin(LightningElement) {
    @track appointments = [];
    @track debugInfo = {};
    @track calendarDays = [];
    @track selectedAppointment = null;
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
    isDesktopFormFactor = FORM_FACTOR === 'Large';
    isCalendarTabActive = false;

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

    // Reschedule with another tech modal
    isRescheduleModalOpen = false;
    rescheduleOptions = [];
    rescheduleSelection = null;
    rescheduleWorkOrderId = null;
    rescheduleLoading = false;

    // Transfer request rejection modal
    isRejectModalOpen = false;
    rejectReason = '';
    rejectRequestId = null;

    // Long press to start drag
    dragLongPressTimer = null;
    isPressingForDrag = false;
    _pendingDrag = null;          // holds data until long press triggers

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

    // swipe-to-close support
    touchStartY = null;

    // ======= GETTERS =======

    get hasAppointments() {
        return this.appointments && this.appointments.length > 0;
    }

    get isMyMode() {
        return this.listMode === 'my';
    }

    get isDragGhostVisible() {
        return this.dragGhostVisible;
    }

    get isDragging() {
        return !!(this.dragMode || this.isPressingForDrag);
    }

    get calendarDaysWrapperClass() {
        return this.isDragging
            ? 'sfs-calendar-days-wrapper sfs-calendar-days-wrapper_dragging'
            : 'sfs-calendar-days-wrapper';
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

    get managerResetDisabled() {
        return !this.isViewingAsOther;
    }

    get rescheduleSubmitDisabled() {
        return this.rescheduleLoading || !this.rescheduleSelection;
    }

    // Position + size of the floating event
    get dragGhostStyle() {
        return `top:${this.dragGhostY}px;left:${this.dragGhostX}px;width:${this.dragGhostWidth}px;height:${this.dragGhostHeight}px;transform:translateX(-50%);`;
    }


    // Classes for the inner event block
    get dragGhostClass() {
        const base = 'sfs-calendar-event sfs-calendar-event_ghost';
        return this.dragGhostTypeClass
            ? `${base} ${this.dragGhostTypeClass}`
            : base;
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
        return this.ownedAppointments.filter(appt =>
            this.isQuoteStatus(appt.workOrderStatus)
        ).length;
    }

    get needQuoteCount() {
        return this.ownedAppointments.filter(appt =>
            appt.workOrderStatus === 'Need Quote'
        ).length;
    }

    get poRequestedCount() {
        return this.ownedAppointments.filter(appt =>
            appt.workOrderStatus === 'PO Requested'
        ).length;
    }

    get quoteSentCount() {
        return this.ownedAppointments.filter(appt =>
            appt.workOrderStatus === 'Quote Sent'
        ).length;
    }

    get quoteAttachedCount() {
        return this.ownedAppointments.filter(appt =>
            this.isQuoteAttachedAppointment(appt)
        ).length;
    }

    get visibleAppointments() {
        if (!this.appointments) return [];

        let baseList = [];

        switch (this.listMode) {
            case 'crew':
                baseList = this.appointments.filter(a => a.isCrewAssignment);
                break;

            case 'partsReady':
                baseList = this.appointments.filter(a => a.allPartsEnRoute);
                break;

            case 'fulfilling':
                baseList = this.appointments.filter(
                    a => a.somePartsEnRoute && !a.allPartsEnRoute
                );
                break;

            case 'quotes':
                baseList = this.ownedAppointments.filter(appt =>
                    this.isQuoteStatus(appt.workOrderStatus)
                );
                break;

            case 'needQuote':
                baseList = this.ownedAppointments.filter(
                    appt => appt.workOrderStatus === 'Need Quote'
                );
                break;

            case 'poRequested':
                baseList = this.ownedAppointments.filter(
                    appt => appt.workOrderStatus === 'PO Requested'
                );
                break;

            case 'quoteSent':
                baseList = this.ownedAppointments.filter(
                    appt => appt.workOrderStatus === 'Quote Sent'
                );
                break;

            case 'quoteAttached':
                baseList = this.ownedAppointments.filter(appt =>
                    this.isQuoteAttachedAppointment(appt)
                );
                break;

            case 'my':
            default: {
                baseList = this.ownedAppointments;
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

        return baseList;
    }

    isQuoteStatus(status) {
        return this.quoteStatuses.includes(status);
    }

    isQuoteAttachedAppointment(appt) {
        const status = appt.workOrderStatus;
        const hasAttachment = appt.hasQuoteAttachment || Boolean(
            appt.quoteAttachmentUrl
        );

        return status === 'Quote Attached' ||
            (status === 'Need Quote' && hasAttachment);
    }

    get hasSelectedAppointment() {
        return this.selectedAppointment !== null;
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

    get listMyModeClass() {
        return this.listMode === 'my'
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isQuotesMode() {
        return this.listMode === 'quotes';
    }

    get listQuotesModeClass() {
        return this.isQuotesMode
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isNeedQuoteMode() {
        return this.listMode === 'needQuote';
    }

    get listNeedQuoteModeClass() {
        return this.isNeedQuoteMode
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isPoRequestedMode() {
        return this.listMode === 'poRequested';
    }

    get listPoRequestedModeClass() {
        return this.isPoRequestedMode
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isQuoteSentMode() {
        return this.listMode === 'quoteSent';
    }

    get listQuoteSentModeClass() {
        return this.isQuoteSentMode
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isQuoteAttachedMode() {
        return this.listMode === 'quoteAttached';
    }

    get listQuoteAttachedModeClass() {
        return this.isQuoteAttachedMode
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get isCrewMode() {
        return this.listMode === 'crew';
    }

    get listCrewModeClass() {
        let classes = 'sfs-mode-btn';

        if (this.isCrewCountUrgent) {
            classes += ' sfs-mode-btn_alert';
        }
        if (this.listMode === 'crew') {
            classes += ' sfs-mode-btn_active';
        }
        return classes;
    }

    get isTransferMode() {
        return this.listMode === 'transferRequests';
    }

    get listTransferModeClass() {
        let classes = 'sfs-mode-btn';

        if (this.transferRequestCount > 0) {
            classes += ' sfs-mode-btn_transfer-alert';
        }

        if (this.isTransferMode) {
            classes += ' sfs-mode-btn_active';
        }

        return classes;
    }

    get listPartsReadyModeClass() {
        return this.listMode === 'partsReady'
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
    }

    get listFulfillingModeClass() {
        return this.listMode === 'fulfilling'
            ? 'sfs-mode-btn sfs-mode-btn_active'
            : 'sfs-mode-btn';
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

    // Tray helpers
    get trayContainerClass() {
        return this.pullTrayOpen ? 'sfs-tray sfs-tray_open' : 'sfs-tray';
    }

    get unscheduledCount() {
        return this.unscheduledWorkOrders
            ? this.unscheduledWorkOrders.length
            : 0;
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

    // ======= LIFECYCLE =======

    connectedCallback() {
        this.checkOnline();
        if (!this.isOffline) {
            this.loadAppointments();
        }
    }

    renderedCallback() {
        if (
            this.isTimelineMode &&
            this._needsCenterOnToday &&
            this.calendarDays &&
            this.calendarDays.length > 0
        ) {
            this.centerTimelineOnTodayColumn();
        }

        this.scheduleNowLinePositionUpdate();
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
        this._needsCenterOnToday = false;
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

    // ======= DRAG HANDLERS (events) =======

    handleEventDragStart(event) {
        // Do not start another drag if one is already running
        if (this.dragMode || this.isPressingForDrag) {
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

        let clientX;
        let clientY;
        if (event.touches && event.touches.length) {
            clientX = event.touches[0].clientX;
            clientY = event.touches[0].clientY;
        } else {
            clientX = event.clientX;
            clientY = event.clientY;
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
            clientX,
            clientY,
            dayBodyHeight: bodyRect.height || dayBodyEl.clientHeight || 1,
            dayBodyTop: bodyRect.top,
            dayWidth: dayEl.clientWidth || 1,
            title: appt.workOrderSubject || 'Appointment'
        };


        this.isPressingForDrag = true;
        this._pendingDrag = pending;

        // Long press threshold (about a quarter second)
        this.clearLongPressTimer();
        this.dragLongPressTimer = window.setTimeout(() => {
            this.beginDragFromPending();
        }, 250);

        event.preventDefault();
        event.stopPropagation();
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
        if (this.dragMode) {
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

        let clientX;
        let clientY;
        if (event.touches && event.touches.length) {
            clientX = event.touches[0].clientX;
            clientY = event.touches[0].clientY;
        } else {
            clientX = event.clientX;
            clientY = event.clientY;
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
        this.dragStartClientX = clientX;
        this.dragStartClientY = clientY;
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
        event.stopPropagation();
        event.preventDefault();
    }



    // ======= DRAG HANDLERS (tray -> calendar) =======
    handleTrayCardDragStart(event) {
        if (this.dragMode || this.isPressingForDrag) {
            return;
        }

        const card = event.currentTarget;
        const workOrderId = card.dataset.woid;
        if (!workOrderId) {
            return;
        }

        let clientX;
        let clientY;
        if (event.touches && event.touches.length) {
            clientX = event.touches[0].clientX;
            clientY = event.touches[0].clientY;
        } else {
            clientX = event.clientX;
            clientY = event.clientY;
        }

        // Use first day column/body to measure width and height
        const dayEl = this.template.querySelector('.sfs-calendar-day');
        const dayBodyEl = this.template.querySelector('.sfs-calendar-day-body');

        if (!dayEl || !dayBodyEl) {
            return;
        }

        const startIndex =
            this.todayDayIndex != null ? this.todayDayIndex : 0;

        const bodyRect = dayBodyEl.getBoundingClientRect();

        const pending = {
            type: 'wo',
            workOrderId,
            dayIndex: startIndex,
            clientX,
            clientY,
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
        this.dragLongPressTimer = window.setTimeout(() => {
            this.beginDragFromPending();
        }, 250);

        event.preventDefault();
        event.stopPropagation();
    }

        clearLongPressTimer() {
        if (this.dragLongPressTimer) {
            window.clearTimeout(this.dragLongPressTimer);
            this.dragLongPressTimer = null;
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

                this.updateSelectedEventStyles();
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




    handleCalendarPointerMove(event) {
        if (!this.dragMode || this.dragStartClientX === null) {
            return;
        }

        let clientX;
        let clientY;
        if (event.touches && event.touches.length) {
            clientX = event.touches[0].clientX;
            clientY = event.touches[0].clientY;
        } else if (event.changedTouches && event.changedTouches.length) {
            clientX = event.changedTouches[0].clientX;
            clientY = event.changedTouches[0].clientY;
        } else {
            clientX = event.clientX;
            clientY = event.clientY;
        }

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
        }

        event.preventDefault();
    }


    handleCalendarPointerEnd(event) {
        if (!this.dragMode || this.dragStartClientX === null) {
            this.resetDragState();
            return;
        }

        let clientY;
        if (event.changedTouches && event.changedTouches.length) {
            clientY = event.changedTouches[0].clientY;
        } else if (event.touches && event.touches.length) {
            clientY = event.touches[0].clientY;
        } else {
            clientY = event.clientY;
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
            this.resetDragState();
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

                this.createAppointmentFromWorkOrder(
                    workOrderId,
                    this.toUserIsoString(dropLocal),
                    this.toUserIsoString(dropEnd)
                );
            }
        }

        this.resetDragState();
        event.preventDefault();
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
        this.updateSelectedEventStyles();
    }



    // ======= DATA LOAD =======

    loadAppointments() {
        this.isLoading = true;
        this.selectedAppointment = null;

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
                    clone.hasQuoteAttachment = Boolean(
                        appt.hasQuoteAttachment ||
                            appt.workOrderStatus === 'Quote Attached')
                    clone.hasQuoteAttachment = Boolean(
                        appt.hasQuoteAttachment
                    );

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

        const offsetMinutes = this.getTimeZoneOffsetMinutes(d);
        const localMs = d.getTime() + offsetMinutes * 60 * 1000;
        return new Date(localMs);
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

        const offsetMinutes = this.getTimeZoneOffsetMinutes(local);
        const utcMs = local.getTime() - offsetMinutes * 60 * 1000;
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
                style: `top:${topPct}%;height:${heightPct}%;`,
                subject: appt.workOrderSubject,
                workOrderNumber: appt.workOrderNumber,
                workTypeName: appt.workTypeName,
                timeLabel,
                isCrewAssignment: appt.isCrewAssignment,
                isMyAssignment: appt.isMyAssignment,
                className: baseTimelineClass,
                classNameWeek: baseWeekClass,
                baseClassTimeline: baseTimelineClass,
                baseClassWeek: baseWeekClass
            });
        });

        if (this.isTimelineMode) {
            this.showNowLine = true;
        }

        this.calendarDays = days;
        this.updateSelectedEventStyles();
        this.scheduleNowLinePositionUpdate();
    }

    scheduleNowLinePositionUpdate() {
        if (this._nowLineFrame) {
            cancelAnimationFrame(this._nowLineFrame);
        }

        this._nowLineFrame = requestAnimationFrame(() => {
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
        const selectedId = this.selectedAppointment
            ? this.selectedAppointment.appointmentId
            : null;

        const draggingId = this.dragMode ? this.draggingEventId : null;

        const updatedDays = this.calendarDays.map(day => {
            const newEvents = day.events.map(evt => {
                const baseTimeline =
                    evt.baseClassTimeline || 'sfs-calendar-event';
                const baseWeek =
                    evt.baseClassWeek || 'sfs-week-event-box';

                const isSelected = selectedId && evt.id === selectedId;
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
        const mode = event.target.dataset.mode;
        if (!mode) {
            return;
        }
        this.listMode = mode;
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
        if (!id) {
            return;
        }
        const appt = this.appointments.find(a => a.appointmentId === id);

        window.clearTimeout(this._closeTimeout);
        this.isDetailClosing = false;
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

        window.clearTimeout(this._closeTimeout);
        this._closeTimeout = window.setTimeout(() => {
            this.selectedAppointment = null;
            this.isDetailClosing = false;
            this.updateSelectedEventStyles();
        }, 200);
    }

    handleOverlayClick() {
        this.handleCloseDetails();
    }

    handleDetailCardClick(event) {
        event.stopPropagation();
    }

    handleDetailTouchStart(event) {
        if (event.touches && event.touches.length > 0) {
            this.touchStartY = event.touches[0].clientY;
        }
    }

    handleDetailTouchEnd(event) {
        if (this.touchStartY === null) {
            return;
        }
        if (event.changedTouches && event.changedTouches.length > 0) {
            const deltaY =
                event.changedTouches[0].clientY - this.touchStartY;
            if (deltaY > 40) {
                this.handleCloseDetails(event);
            }
        }
        this.touchStartY = null;
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

    handleManagerTabActive() {
        this.isCalendarTabActive = false;
        this.pullTrayOpen = false;
    }

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

    handleCalendarTabActive() {
        this.isCalendarTabActive = true;
        this.handleCalendarToday();
    }

    handleListTabActive() {
        this.isCalendarTabActive = false;
        this.pullTrayOpen = false;
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

            window.clearTimeout(this._centerTimeout);
            this._centerTimeout = window.setTimeout(() => {
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

    // ======= TRAY HANDLERS =======

    toggleTray() {
        this.pullTrayOpen = !this.pullTrayOpen;
    }

    handleTrayCardClick(event) {
        // If we started a drag, ignore the click
        if (this.dragHasMoved) {
            this.dragHasMoved = false;
            return;
        }

        this.showToast(
            'Drag to schedule',
            'Drag this card onto a day in the calendar to create a new appointment.',
            'info'
        );
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

        this.checkOnline();
        if (this.isOffline) {
            this.showToast(
                'Offline',
                'You must be online to update an assignment.',
                'warning'
            );
            return;
        }

        const confirmed = confirm(
            'Are you sure you want to unschedule this service appointment?'
        );
        if (!confirmed) {
            return;
        }

        this.isLoading = true;

        unassignAppointment({ appointmentId: id })
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
        if (!workTypeName) {
            return '';
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

    reduceError(error) {
        let message = 'Unknown error';
        if (Array.isArray(error?.body)) {
            message = error.body.map(e => e.message).join(', ');
        } else if (error?.body?.message) {
            message = error.body.message;
        } else if (error?.message) {
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
