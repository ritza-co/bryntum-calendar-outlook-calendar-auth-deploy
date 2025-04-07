import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { createEvent, updateEvent, deleteEvent } from './GraphService';
import { EventModel } from '@bryntum/calendar';
import React from 'react';
import { differenceInDays } from 'date-fns';

export async function BryntumSync(
    id: string,
    name: string,
    startDate: string,
    endDate: string,
    allDay: boolean,
    action: 'remove' | 'removeAll' | 'add' | 'clearchanges' | 'filter' | 'update' | 'dataset' | 'replace',
    setEvents: React.Dispatch<React.SetStateAction<Partial<EventModel>[] | undefined>>,
    authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string
) {
    // For all-day events, set the time to midnight
    const adjustDateForAllDay = (start: string, end: string) => {
        const startDate = new Date(start);
        const endDate = new Date(end);

        // Extract year, month, day directly from the date
        const startYear = startDate.getFullYear();
        const startMonth = startDate.getMonth(); // Already 0-based
        const startDay = startDate.getDate();

        const endYear = endDate.getFullYear();
        const endMonth = endDate.getMonth(); // Already 0-based
        const duration = differenceInDays(endDate, startDate);
        const endDay = endDate.getDate() + (duration === 0 ? 1 : 0);

        const utcStartDate = new Date(Date.UTC(startYear, startMonth, startDay));
        const utcEndDate = new Date(Date.UTC(endYear, endMonth, endDay));

        return { utcStartDate : utcStartDate.toISOString(), utcEndDate : utcEndDate.toISOString() };
    };

    const { utcStartDate, utcEndDate } = adjustDateForAllDay(startDate, endDate);

    const eventData = {
        subject : name || 'New Event',
        start   : {
            dateTime : allDay ? utcStartDate : startDate,
            timeZone
        },
        end : {
            dateTime : allDay ? utcEndDate : endDate,
            timeZone
        },
        isAllDay : allDay
    };

    const formatEventDate = (dateTime: string, eventTimeZone: string, isAllDay: boolean) => {
        if (isAllDay) {
            // For all-day events, just pass through the UTC date
            return dateTime;
        }
        // If the date is already in UTC format, just ensure it has the correct milliseconds format
        if (eventTimeZone === 'UTC') {
            return dateTime.replace(/\.0+Z?$/, '.000Z');
        }
        // For non-UTC dates, convert to UTC
        return new Date(dateTime).toISOString();
    };

    try {
        if (action === 'add') {
            const result = await createEvent(authProvider, eventData);
            if (result.id && result.start?.dateTime && result.end?.dateTime) {
                const newEvent: Partial<EventModel> = {
                    id        : result.id,
                    name      : result.subject || '',
                    startDate : formatEventDate(result.start.dateTime, result.start.timeZone || timeZone, result.isAllDay || false),
                    endDate   : formatEventDate(result.end.dateTime, result.end.timeZone || timeZone, result.isAllDay || false),
                    allDay    : result.isAllDay || false
                };
                setEvents(prev => prev ? [...prev, newEvent] : [newEvent]);
            }
        }
        else if (action === 'update') {
            if (!id || id.startsWith('_generated')) return;
            const result = await updateEvent(authProvider, id, eventData);

            if (result.id && result.start?.dateTime && result.end?.dateTime) {
                const updatedEvent: Partial<EventModel> = {
                    id        : result.id,
                    name      : result.subject || '',
                    startDate : formatEventDate(result.start.dateTime, result.start.timeZone || timeZone, result.isAllDay || false),
                    endDate   : formatEventDate(result.end.dateTime, result.end.timeZone || timeZone, result.isAllDay || false),
                    allDay    : result.isAllDay || false
                };
                setEvents(prevEvents =>
                    prevEvents?.map(evt => (evt.id === id ? updatedEvent : evt))
                );
            }
        }
        else if (action === 'remove') {
            if (!id || id.startsWith('_generated')) return;
            await deleteEvent(authProvider, id);
            setEvents(prevEvents => prevEvents?.filter(evt => evt.id !== id));
        }
    }
    catch (error) {
        console.error('Error syncing with Outlook calendar:', error);
        throw error;
    }
}

