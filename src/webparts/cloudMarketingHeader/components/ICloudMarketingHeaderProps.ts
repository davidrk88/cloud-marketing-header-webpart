import { ITargetAudienceProps } from '../../../common/TargetAudience';
export interface ICloudMarketingHeaderProps extends ITargetAudienceProps {
	// bgimage: string;
	filePickerResult: any;
	gradientcolor: string;
	title: string;
	description: string;
	descriptionat: string;
	textcolor: string;
	btncolor: string;
	linktxtcolor: string;
	showconfidential: boolean;
	showannouncement: boolean;
	announcementmsg: string;
	announcementtype: string;
	showannouncelink: boolean;
	announcementlink: string;
	announcementlinktext: string;
	showhelplink: boolean;
	helplinktext: string;
	helplinkurl: string;
	helplinkinvert: boolean;
	context: any;
	audienceTargetsDesc: any[];
	headerLinksConfig: any[];
}
