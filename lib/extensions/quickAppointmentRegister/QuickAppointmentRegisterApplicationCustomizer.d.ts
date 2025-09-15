import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
export interface IQuickAppointmentRegisterApplicationCustomizerProperties {
}
export default class QuickAppointmentRegisterApplicationCustomizer extends BaseApplicationCustomizer<IQuickAppointmentRegisterApplicationCustomizerProperties> {
    onInit(): Promise<void>;
    private specialClientSideExtensions;
    private extendEventPage;
    private loadAppointment;
    private manageUserToAppointment;
}
//# sourceMappingURL=QuickAppointmentRegisterApplicationCustomizer.d.ts.map