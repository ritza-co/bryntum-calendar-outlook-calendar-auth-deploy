import { useAppContext } from '../AppContext';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { findIana } from 'windows-iana';
import { getUserFutureCalendar, getUserPastCalendar, getUserWeekCalendar } from '../GraphService';
import SignInModal from './SignInModal';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { BryntumButton, BryntumCalendar } from '@bryntum/calendar-react';
import bryntumLogo from '../assets/bryntum-symbol-white.svg';
import { EventModel, Model, ProjectConsumer, ProjectModelMixin, Store, Toast } from '@bryntum/calendar';
import { BryntumSync } from '../crudHelpers';

interface SyncDataParams {
    source: typeof ProjectConsumer;
    project: typeof ProjectModelMixin;
    store: Store;
    action: 'remove' | 'removeAll' | 'add' | 'clearchanges' | 'filter' | 'update' | 'dataset' | 'replace';
    record: Model;
    records: Model[];
    changes: object
}

export default function Calendar() {
    const app = useAppContext();
    const [events, setEvents] = useState<Partial<EventModel>[]>();
    const calendarRef = useRef(null);
    const hasFetchedInitialEvents = useRef(false);
    const hasRunFirstEffect = useRef(false);

    // Convert Windows timezone to IANA format for Bryntum
    const timeZone = useMemo(() => {
        if (!app.user?.timeZone) return 'UTC';
        const ianaTimeZones = findIana(app.user.timeZone);
        return ianaTimeZones[0].valueOf();
    }, [app.user?.timeZone]);

    const syncWithOutlook = useCallback((
        id: string,
        name: string,
        startDate: string,
        endDate: string,
        allDay: boolean,
        action: 'add' | 'update' | 'remove' | 'removeAll' | 'clearchanges' | 'filter' | 'update' | 'dataset' | 'replace'
    ) => {
        if (!app.authProvider) return;
        return BryntumSync(
            id,
            name,
            startDate,
            endDate,
            allDay,
            action,
            setEvents,
            app.authProvider,
            timeZone
        );
    }, [app.authProvider, timeZone]);

    const syncData = useCallback(({ action, records }: SyncDataParams) => {
        if ((action === 'add' && !records[0].copyOf) || action === 'dataset') {
            return;
        }
        records.forEach((record) => {
            syncWithOutlook(
                record.get('id'),
                record.get('name'),
                record.get('startDate'),
                record.get('endDate'),
                record.get('allDay'),
                action
            );
        });
    }, [syncWithOutlook]);

    const addRecord = useCallback((event: { eventRecord: EventModel }) => {
        const { eventRecord } = event;
        const isNew = eventRecord.id.toString().startsWith('_generated');

        // Get the date strings in the calendar's timezone
        const startDate = new Date(eventRecord.startDate);
        const endDate = new Date(eventRecord.endDate);

        syncWithOutlook(
            eventRecord.id.toString(),
            eventRecord.name || '',
            startDate.toISOString(),
            endDate.toISOString(),
            eventRecord.allDay || false,
            isNew ? 'add' : 'update'
        );
    }, [syncWithOutlook]);

    useEffect(() => {
        const loadEvents = async() => {
            if (app.user && !events) {
                if (hasRunFirstEffect.current) {
                    return;
                }
                hasRunFirstEffect.current = true;
                try {
                    const ianaTimeZones = findIana(app.user?.timeZone || 'UTC');
                    const outlookEvents = await getUserWeekCalendar(app.authProvider!, ianaTimeZones[0].valueOf());
                    const calendarEvents: Partial<EventModel>[] = [];
                    outlookEvents.forEach((event) => {
                        // Convert the dates to the calendar's timezone
                        const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                        const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                        calendarEvents.push({
                            id        : `${event.id}`,
                            name      : `${event.subject}`,
                            startDate : startDate?.toISOString(),
                            endDate   : endDate?.toISOString(),
                            allDay    : event.isAllDay || false
                        });
                    });
                    setEvents(calendarEvents);
                }
                catch (err) {
                    const error = err as Error;
          app.displayError!(error.message);
                }
            }
        };

        loadEvents();
    }, [app.user, app.authProvider, events, app.displayError]);

    // load more events from outlook
    useEffect(() => {
        // Skip if still loading or if we've already fetched events
        if (!app.user) return;
        if (app.isLoading) return;
        if (hasFetchedInitialEvents.current) return;

        async function fetchAllData() {
            try {
                hasFetchedInitialEvents.current = true;
                const ianaTimeZones = findIana(app.user?.timeZone || 'UTC');
                const [calendarPastEvents, calendarFutureEvents] = await Promise.all([getUserPastCalendar(app.authProvider!, ianaTimeZones[0].valueOf()), getUserFutureCalendar(app.authProvider!, ianaTimeZones[0].valueOf())]);

                const pastEvents: Partial<EventModel>[] = [];
                const futureEvents: Partial<EventModel>[] = [];

                calendarPastEvents.forEach((event) => {
                    // Convert the dates to the calendar's timezone
                    const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                    const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                    pastEvents.push({
                        id        : `${event.id}`,
                        name      : `${event.subject}`,
                        startDate : startDate?.toISOString(),
                        endDate   : endDate?.toISOString(),
                        allDay    : event.isAllDay || false
                    });
                });

                calendarFutureEvents.forEach((event) => {
                    // Convert the dates to the calendar's timezone
                    const startDate = event.start?.dateTime ? new Date(event.start.dateTime) : null;
                    const endDate = event.end?.dateTime ? new Date(event.end.dateTime) : null;

                    futureEvents.push({
                        id        : `${event.id}`,
                        name      : `${event.subject}`,
                        startDate : startDate?.toISOString(),
                        endDate   : endDate?.toISOString(),
                        allDay    : event.isAllDay || false
                    });
                });

                setEvents(currentEvents => {
                    return [...pastEvents, ...(currentEvents || []), ...futureEvents];
                });
            }
            catch (err) {
                const error = err as Error;
                console.log('dfd');

                app.displayError!(error.message);
            }
        }
        fetchAllData();
    }, [app.user, app.authProvider, app.displayError, app.isLoading, app.user?.timeZone]);


    useEffect(() => {
        if (app.error) {
            Toast.show({
                html    : app.error.message,
                timeout : 0
            });
        }
    }, [app.error]);

    const calendarConfig = useMemo(() => ({
        defaultMode      : 'week',
        timeZone         : 'Africa/Johannesburg',
        eventEditFeature : {
            items : {
                nameField : {
                    required : true
                },
                resourceField   : null,
                recurrenceCombo : null
            }
        },
        eventMenuFeature : {
            items : {
                duplicate : null
            }
        },
        onDataChange     : syncData,
        onAfterEventSave : addRecord
    }), [syncData, addRecord]);

    return (
        <>
            <UnauthenticatedTemplate>
                <SignInModal />
            </UnauthenticatedTemplate>
            <header>
                <div className="title-container">
                    <img src={bryntumLogo} role="presentation" alt="Bryntum logo" />
                    <h1>Bryntum Calendar synced with Outlook Calendar demo</h1>
                </div>
                <AuthenticatedTemplate>
                    <BryntumButton
                        cls="b-raised"
                        text={app.user && app.isLoading ? 'Signing out...' : 'Sign out'}
                        color='b-blue'
                        onClick={() => app.signOut?.()}
                        disabled={app.isLoading}
                    />
                </AuthenticatedTemplate>
                <UnauthenticatedTemplate>
                    <BryntumButton
                        cls="b-raised"
                        text={app.isLoading ? 'Signing in...' : 'Sign in with Microsoft'}
                        color='b-blue'
                        onClick={() => app.signIn?.()}
                        disabled={app.isLoading}
                    />
                </UnauthenticatedTemplate>
            </header>
            <BryntumCalendar
                ref={calendarRef}
                eventStore={{
                    data : events
                }}
                {...calendarConfig}
            />
            <div className="notice-info">
                <div className="notice-info-container">
                    <p className='notice-info-text'>
                This is a publicly accessible demonstration of the{' '}
                        <a href="https://bryntum.com/products/calendar/">
                      Bryntum Calendar Component.
                        </a>{' '}
                    By signing in, you&apos;ll be able to see how real events from your
                    Outlook Calendar are displayed in the Bryntum Calendar. You&apos;ll also
                    be able to edit your events in the Bryntum Calendar component and
                    see those changles reflect on your Outlook Calendar. Note that after
                    signing in, you&apos;ll need to grant us read and write access to your
                    Outlook Calendar. We do not store any events for longer than needed
                    to display them and send any changes back to Outlook Calendar and we
                    do not do anything with your data beyond what is strictly needed for
                    the demonstration.
                    </p>
                    <p className='notice-info-links'>
                To provide these features, we request access to your Outlook Calendar
                data. View our
                        <a href="/privacypolicy.html"> Privacy Policy</a> and{' '}
                        <a href="/termsofservice.html">Terms of Service</a>.
                    </p>
                </div>
            </div>
        </>
    );
}
