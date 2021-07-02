import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneDropdown,
	PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { 
	PropertyFieldCollectionData, 
	CustomCollectionFieldType 
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { 
	PropertyFieldFilePicker, 
	IPropertyFieldFilePickerProps, 
	IFilePickerResult 
} from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { sp } from "@pnp/sp";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CloudMarketingHeaderWebPartStrings';
import CloudMarketingHeader from './components/CloudMarketingHeader';
import { ICloudMarketingHeaderProps } from './components/ICloudMarketingHeaderProps';

// Telemetry opt-out for PnP controls - required for submission to MSIT
import PnPTelemetry from "@pnp/telemetry-js";
const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

export interface ICloudMarketingHeaderWebPartProps {
	filePickerResult: IFilePickerResult;
	gradientcolor: string;
	title: string;
	description: string;
	textcolor: string;
	btncolor: string;
	linktxtcolor: string;
	showannouncement: boolean;
	announcementmsg: string;
	announcementtype: string;
	context: any;
	headerLinksConfig: any[];
}

export default class CloudMarketingHeaderWebPart extends BaseClientSideWebPart<ICloudMarketingHeaderWebPartProps> {

	public onInit(): Promise<void> {
		return super.onInit().then(_ => {
			sp.setup({
				spfxContext: this.context
			});
		});
	}

	public render(): void {
		let bgImgUrl;
		let gradColor;

		const element: React.ReactElement<ICloudMarketingHeaderProps> = React.createElement(
			CloudMarketingHeader,
			{
				filePickerResult: this.properties.filePickerResult,
				gradientcolor: this.properties.gradientcolor,
				title: this.properties.title,
				description: this.properties.description,
				textcolor: this.properties.textcolor,
				btncolor: this.properties.btncolor,
				linktxtcolor: this.properties.linktxtcolor,
				showannouncement: this.properties.showannouncement,
				announcementmsg: this.properties.announcementmsg,
				announcementtype: this.properties.announcementtype,
				context: this.context,
				headerLinksConfig: this.properties.headerLinksConfig
			}
		);


		if(this.properties.gradientcolor == "bgColorWhite") {
			gradColor = "rgba(255,255,255,1), rgba(255,255,255,0)";
		} else {
			gradColor = "rgba(0,0,0,1), rgba(0,0,0,0)";
		}
		
		// bgImgUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/${this.properties.bgimage}`;
		bgImgUrl = "";
		if(this.properties.filePickerResult !== undefined) {
			bgImgUrl = `${this.properties.filePickerResult.fileAbsoluteUrl}`;
		}
		ReactDom.render(element, this.domElement);

		this.domElement.querySelector('#cloudMarketingHeaderWebpartMain').setAttribute("style", "background-image: linear-gradient(to right, " + gradColor + "), url('" + bgImgUrl + "');");
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					displayGroupsAsAccordion: true,
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							isCollapsed: true,
							groupFields: [
								PropertyFieldFilePicker('filePicker', {
									context: this.context,
									filePickerResult: this.properties.filePickerResult,
									onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
									properties: this.properties,
									onSave: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e;  },
									onChanged: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e; },
									key: "filePickerId",
									buttonLabel: strings.BGImgButtonLabel,
									label: strings.BGImgSelectLabel,
									buttonIcon: "DocumentSearch",
									accepts: [".jpg", "jpeg", ".png", ".svg"]
								}),
								PropertyPaneDropdown('gradientcolor', {
									label: strings.GradientColorFieldLabel,
									options: [
										{ key: "bgColorWhite", text: "White" },
										{ key: "bgColorBlack", text: "Black" }
									],
									selectedKey: "bgColorBlack"
								}),
								PropertyPaneTextField('title', {
									label: strings.TitleFieldLabel
								}),
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
									multiline: true
								})
							]
						},
						{
							groupName: strings.LinksGroupName,
							isCollapsed: true,
							groupFields: [
								PropertyFieldCollectionData("headerLinksConfig", {
									key: "headerLinksConfig",
									label: "", //"Header Links Settings",
									panelHeader: strings.LinksPanelHeader,
									manageBtnLabel: strings.LinksPanelBtnLabel,
									value: this.properties.headerLinksConfig,
									fields: [
										{
											id: "linkText",
											title: "Display Text",
											placeholder: "Enter text here...",
											type: CustomCollectionFieldType.string,
											required: true
										},
										{
											id: "linkUrl",
											title: "URL",
											// defaultValue: "https://",
											placeholder: "https://",
											type: CustomCollectionFieldType.string
										},
										{
											id: "linkOrder",
											title: "Sort Order",
											type: CustomCollectionFieldType.number
										},
										{
											id: "linkFlag",
											title: "Enable",
											defaultValue: true,
											type: CustomCollectionFieldType.boolean
										}
									],
									disabled: false
								}),
								PropertyPaneDropdown('textcolor', {
									label: strings.TextColorFieldLabel,
									options: [
										{ key: "txtColorWhite", text: "White" },
										{ key: "txtColorBlack", text: "Black" },
										{ key: "txtColorRed", text: "Red" },
										{ key: "txtColorOrange", text: "Orange" },
										{ key: "txtColorYellow", text: "Yellow" },
										{ key: "txtColorGreen", text: "Green" },
										{ key: "txtColorBlue", text: "Blue" },
										{ key: "txtColorPurple", text: "Purple" },
										{ key: "txtColorPink", text: "Pink" },
										{ key: "txtColorCyan", text: "Cyan" }
									],
									selectedKey: "txtColorWhite"
								}),

								PropertyPaneDropdown('linktxtcolor', {
									label: strings.LinkTxtColorFieldLabel,
									options: [
										{ key: "txtColorWhite", text: "White" },
										{ key: "txtColorBlack", text: "Black" }
									],
									selectedKey: "txtColorBlack"
								}),
								PropertyPaneDropdown('btncolor', {
									label: strings.BTNColorFieldLabel,
									options: [
										{ key: "bgColorDefault", text: "Theme default" },
										{ key: "bgColorWhite", text: "White" },
										{ key: "bgColorBlack", text: "Black" },
										{ key: "bgColorRed", text: "Red" },
										{ key: "bgColorOrange", text: "Orange" },
										{ key: "bgColorYellow", text: "Yellow" },
										{ key: "bgColorGreen", text: "Green" },
										{ key: "bgColorBlue", text: "Blue" },
										{ key: "bgColorPurple", text: "Purple" },
										{ key: "bgColorPink", text: "Pink" },
										{ key: "bgColorCyan", text: "Cyan" }
									],
									selectedKey: "bgColorWhite"
								})
							]
						},
						{
							groupName: strings.AnnouncementGroupName,
							isCollapsed: true,
							groupFields: [
								PropertyPaneToggle('showannouncement', {
									label: strings.AnnouncementFieldLabel,
									checked: false,
									onText: "Enable",
									offText: "Disable"
								}),
								PropertyPaneTextField('announcementmsg', {
									label: strings.AnnouncementMsgFieldLabel,
									placeholder: strings.AnnouncementMsgFieldHolder,
									multiline: true
								}),
								PropertyPaneDropdown('announcementtype', {
									label: strings.AnnouncementTypeFieldLabel,
									options: [
										{ key: "colorInformational", text: "Informational" },
										{ key: "colorWarning", text: "Warning" },
										{ key: "colorCritical", text: "Critical" }
									],
									selectedKey: "colorInformational"
								})
							]
						}
					]
				}
			]
		};
	}
}


