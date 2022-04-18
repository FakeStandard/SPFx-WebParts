import { MessageBarType } from "office-ui-fabric-react";

export interface IPeoplepickerSampleStates {
    title: string;
    users: number[];
    managers: number[];
    showMessageBar: boolean;
    messageType?: MessageBarType;
    message?: string;
}
