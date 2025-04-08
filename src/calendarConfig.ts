import { BryntumCalendarProps } from '@bryntum/calendar-react';
import { SyncDataParams } from './types';
import { EventModel } from '@bryntum/calendar';

type AddRecordParams = {
  eventRecord: EventModel;
}
 type CalendarConfigOptions = {
    syncData: (param: SyncDataParams) => void;
   addRecord: (event: AddRecordParams) => void;
  };

export function createCalendarConfig({ syncData, addRecord }: CalendarConfigOptions): BryntumCalendarProps {
    return {
        mode             : 'week',
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
        onDataChange     : syncData,
        onAfterEventSave : addRecord
    };
}