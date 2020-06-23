import * as React from "react";
import { sp } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import { IEventsWebPartProps } from "./IEventsWebPartProps";
import styles from "./EventsPageComponent.module.sass";

class EventsPageComponent extends React.Component<IEventsWebPartProps, any>{
    constructor(props) {
        super(props);
        this.state = { events: null }
    }

    async componentDidMount() {
        pnp.setup({
            spfxContext: this.props.context
        });

        const xml = `
            <Query>\
                <Where>\
                    <DateRangesOverlap>\
                        <FieldRef Name="EventDate"></FieldRef>\
                        <FieldRef Name="EndDate"></FieldRef>\
                        <FieldRef Name="RecurrenceID"></FieldRef>\
                        <Value Type="DateTime">\
                            <Now />\
                        </Value>\
                    </DateRangesOverlap>\
                </Where>\
            </Query>`;

        sp.web.lists.getByTitle('SampleEvents')
            .renderListDataAsStream({
                OverrideViewXml: xml
            })
            .then(result => {
                this.setState({ events: result });
            })
            .catch(console.log);
    }

    getMonth(date: string) {
        let months = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"];
        let monthIndex = date.indexOf("/");
        let month: number = parseInt(date.substring(0, monthIndex));
        return months[month - 1].toUpperCase();
    }

    getDate(date: string) {
        let dateStartIndex = date.indexOf("/");
        let dateEndIndex = date.lastIndexOf("/");
        return date.substring(dateStartIndex + 1, dateEndIndex);
    }

    getTime(date: string) {
        let timeIndex = date.indexOf(" ");
        return date.substring(timeIndex);
    }

    isSingleDayEvent(event: { EventDate: string; EndDate: string; }) {
        let startDate = event.EventDate;
        let endDate = event.EndDate;
        let startDateMonth = parseInt(startDate.substring(0, startDate.indexOf("/")));
        let endDateMonth = parseInt(endDate.substring(0, endDate.indexOf("/")));
        if (startDateMonth != endDateMonth) return false;
        else {
            let startDateDay = parseInt(this.getDate(startDate));
            let endDateDay = parseInt(this.getDate(endDate));
            if (startDateDay != endDateDay) return false;
            else {
                let startDateYear = parseInt(startDate.substring(startDate.lastIndexOf("/") + 1, startDate.indexOf(" ")));
                let endDateYear = parseInt(endDate.substring(endDate.lastIndexOf("/") + 1, endDate.indexOf(" ")));
                if (startDateYear != endDateYear) return false;
            }
        }
        return true;
    }

    render() {
        let eventsList = undefined;
        if (this.state.events != null) {
            eventsList = this.state.events.Row.map((event) => {
                return (
                    <div className={styles.eventRow}>
                        <div className={styles.date}>
                            <div className={styles.text}>{this.getMonth(event.EventDate)}</div>
                            <div className={styles.text}>{this.getDate(event.EventDate)}</div>
                        </div>
                        <div className={styles.eventDetails}>
                            <div className={styles.title}>{event.Title}</div>
                            {event.fAllDayEvent == "Yes" ?
                                <div>All Day Event, {event.Location}</div> :
                                this.isSingleDayEvent(event) ? <div>{this.getTime(event.EventDate)} - {this.getTime(event.EndDate)}, {event.Location} </div>
                                    : <div>{event.EventDate} - {event.EndDate} </div>
                            }
                        </div>
                    </div>
                )
            })
        }
        return (
            <div>
                {this.state.events != null ? eventsList : ""}</div>
        )
    }
}

export default EventsPageComponent;

