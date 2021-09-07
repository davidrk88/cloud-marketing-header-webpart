import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneDropdown,
	PropertyPaneToggle,
	PropertyPaneLink
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
import { 
	PropertyFieldColorPicker,
	PropertyFieldColorPickerStyle 
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { 
	PropertyFieldPeoplePicker, 
	PrincipalType as PrincipalTypeProp,
	IPropertyFieldGroupOrPerson
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { 
	PeoplePicker, 
	PrincipalType 
} from '@pnp/spfx-controls-react/lib/controls/peoplepicker';
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
	descriptionAT: string;
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
	headerLinksConfig: any[];
	audienceTargetsDesc: IPropertyFieldGroupOrPerson[];
	audienceTargetsLinks: any[];
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
		let bgImgUrl = "";
		let gradColor;

		const element: React.ReactElement<ICloudMarketingHeaderProps> = React.createElement(
			CloudMarketingHeader,
			{
				filePickerResult: this.properties.filePickerResult,
				gradientcolor: this.properties.gradientcolor,
				title: this.properties.title,
				description: this.properties.description,
				descriptionat: this.properties.descriptionAT,
				textcolor: this.properties.textcolor,
				btncolor: this.properties.btncolor,
				linktxtcolor: this.properties.linktxtcolor,
				showconfidential: this.properties.showconfidential,
				showannouncement: this.properties.showannouncement,
				announcementmsg: this.properties.announcementmsg,
				announcementtype: this.properties.announcementtype,
				showannouncelink: this.properties.showannouncelink,
				announcementlink: this.properties.announcementlink,
				announcementlinktext: this.properties.announcementlinktext,
				showhelplink: this.properties.showhelplink,
				helplinktext: this.properties.helplinktext,
				helplinkurl: this.properties.helplinkurl,
				helplinkinvert: this.properties.helplinkinvert,
				context: this.context,
				headerLinksConfig: this.properties.headerLinksConfig,
				audienceTargetsDesc: this.properties.audienceTargetsDesc,
				audienceTargets: this.properties.audienceTargetsLinks,
				pageContext: this.context.pageContext
			}
		);

		if(this.properties.gradientcolor == "bgColorWhite") {
			gradColor = "rgba(255,255,255,1), rgba(255,255,255,0)";
		} else {
			gradColor = "rgba(0,0,0,1), rgba(0,0,0,0)";
		}
		if(this.properties.filePickerResult !== undefined) {
			bgImgUrl = `${this.properties.filePickerResult.fileAbsoluteUrl}`;
		}

		ReactDom.render(element, this.domElement);

		this.domElement.querySelector('#cloudMarketingHeaderWebpartMain').setAttribute("style", "background-image: linear-gradient(to right, " + gradColor + "), url('" + bgImgUrl + "');");
		// console.log(this.properties.headerLinksConfig);
		// console.log(this.properties.audienceTarget);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		let announcementMsgControl: any = [];
		let announcementTypeControl: any = [];
		let announceLinkToggleControl: any = [];
		let announceLinkControl: any = [];
		let announceLinkTextControl: any = [];
		let helplinkTextControl: any = [];
		let helplinkURLControl: any = [];
		let helpLinkInvertControl: any = [];

		if(this.properties.showannouncement) {
			announcementMsgControl = PropertyPaneTextField('announcementmsg', {
				label: strings.AnnouncementMsgFieldLabel,
				placeholder: strings.AnnouncementMsgFieldHolder,
				multiline: true
			});

			announcementTypeControl = PropertyPaneDropdown('announcementtype', {
				label: strings.AnnouncementTypeFieldLabel,
				options: [
					{ key: "colorInformational", text: "Informational" },
					{ key: "colorWarning", text: "Warning" },
					{ key: "colorCritical", text: "Critical" }
				],
				selectedKey: "colorInformational"
			});

			announceLinkToggleControl = PropertyPaneToggle('showannouncelink', {
				label: strings.AnnouncementLinkToggle,
				checked: false,
				onText: "Enable",
				offText: "Disable"
			});

			if(this.properties.showannouncelink) {
				announceLinkControl = PropertyPaneTextField('announcementlink', {
					label: strings.AnnouncementLinkLabel,
					value: "https://"
				});

				announceLinkTextControl = PropertyPaneTextField('announcementlinktext', {
					label: strings.AnnouncementLinkTextLabel,
					placeholder: strings.AnnouncementLinkTextPlaceholder
				});
			}
		}

		if(this.properties.showhelplink) {
			helplinkTextControl = PropertyPaneTextField('helplinktext', {
				label: strings.HelpLinkTextLabel,
				value: strings.HelpLinkTextPlaceholder
			});

			helplinkURLControl = PropertyPaneTextField('helplinkurl', {
				label: strings.HelpLinkURLLabel,
				value: "https://"
			});

			helpLinkInvertControl = PropertyPaneToggle('helplinkinvert', {
				label: strings.HelpLinkInvertLabel,
				checked: true,
				onText: "Inverted",
				offText: "Normal" 
			});
		}

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
									// onSave: (e: IFilePickerResult) => { 
									// 	console.log(e); 
									// 	sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl+"/SiteAssets")
									// 		.files.add(e.name, e, true)
									// 		.then((data) => {
									// 			console.log("Upload successful");
									// 			console.log("TEST DATA: " + data);
									// 			this.properties.filePickerResult = data;
									// 		})
									// 		.catch((error) => {
									// 			console.log("Upload failed");
									// 			console.log(error);
									// 		});
									// },
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
								}),
								PropertyPaneTextField('descriptionAT', {
									label: strings.DescriptionATFieldLabel,
									multiline: true
								}),
								PropertyFieldPeoplePicker('audienceTargetsDesc', {
									label: strings.DescriptionATPickerLabel,
									initialData: this.properties.audienceTargetsDesc,
									allowDuplicate: false,
									principalType: [PrincipalTypeProp.SharePoint],
									onPropertyChange: this.onPropertyPaneFieldChanged,
									context: this.context,
									properties: this.properties,
									onGetErrorMessage: null,
									deferredValidationTime: 0,
									key: 'peopleFieldId'
								}),
							]
						},
						{
							groupName: strings.LinksGroupName,
							isCollapsed: true,
							groupFields: [
								PropertyFieldCollectionData("headerLinksConfig", {
									key: "headerLinksConfig",
									label: "", //"Header Links Settings",
									enableSorting: true,
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
											defaultValue: "https://",
											// placeholder: "https://",
											type: CustomCollectionFieldType.string
										},
										{
											id: "audienceTargetsLinks",
											title: "People/Group Picker",
											type: CustomCollectionFieldType.custom,
											onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
												return (  
													React.createElement(PeoplePicker, {  
														context: this.context,
														personSelectionLimit: 10,
														showtooltip: true,
														key: itemId,
														defaultSelectedUsers: item.audienceTargetsLinks,  
														onChange: (items: any[]) => {  
															console.log('Items:', items);
															let tempTargets = [];
															items.map((target) => {
																tempTargets.push(target.secondaryText);
															});
															console.log('tempTargets:',tempTargets);
															item.audienceTargetsLinks = tempTargets;
															onUpdate(field.id, item.audienceTargetsLinks);
														},
														showHiddenInUI: false,  
														principalTypes: [PrincipalType.User, PrincipalType.SharePointGroup]  
													})
												);
											}
										},
										{
											id:"subMenuFlag",
											title: "Link Sub-menu",
											defaultValue: false,
											type: CustomCollectionFieldType.boolean
										},
										// {
										// 	id: "linkOrder",
										// 	title: "Sort Order",
										// 	type: CustomCollectionFieldType.number
										// },
										{
											id: "linkFlag",
											title: "Enable",
											defaultValue: true,
											type: CustomCollectionFieldType.boolean
										}
									],
									disabled: false
								}),
								// PropertyPaneDropdown('textcolor', {
								// 	label: strings.TextColorFieldLabel,
								// 	options: [
								// 		{ key: "txtColorWhite", text: "White" },
								// 		{ key: "txtColorBlack", text: "Black" },
								// 		{ key: "txtColorRed", text: "Red" },
								// 		{ key: "txtColorOrange", text: "Orange" },
								// 		{ key: "txtColorYellow", text: "Yellow" },
								// 		{ key: "txtColorGreen", text: "Green" },
								// 		{ key: "txtColorBlue", text: "Blue" },
								// 		{ key: "txtColorPurple", text: "Purple" },
								// 		{ key: "txtColorPink", text: "Pink" },
								// 		{ key: "txtColorCyan", text: "Cyan" }
								// 	],
								// 	selectedKey: "txtColorWhite"
								// }),
								PropertyFieldColorPicker('textcolor', {
									label: strings.TextColorFieldLabel,
									selectedColor: this.properties.textcolor,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									disabled: false,
									// debounce: 1000,
									isHidden: false,
									alphaSliderHidden: false,
									style: PropertyFieldColorPickerStyle.Full,
									iconName: 'Precipitation',
									key: 'textcolorFieldId'
								}),
								// PropertyPaneDropdown('btncolor', {
								// 	label: strings.BTNColorFieldLabel,
								// 	options: [
								// 		{ key: "bgColorDefault", text: "Theme default" },
								// 		{ key: "bgColorWhite", text: "White" },
								// 		{ key: "bgColorBlack", text: "Black" },
								// 		{ key: "bgColorRed", text: "Red" },
								// 		{ key: "bgColorOrange", text: "Orange" },
								// 		{ key: "bgColorYellow", text: "Yellow" },
								// 		{ key: "bgColorGreen", text: "Green" },
								// 		{ key: "bgColorBlue", text: "Blue" },
								// 		{ key: "bgColorPurple", text: "Purple" },
								// 		{ key: "bgColorPink", text: "Pink" },
								// 		{ key: "bgColorCyan", text: "Cyan" }
								// 	],
								// 	selectedKey: "bgColorWhite"
								// })
								PropertyFieldColorPicker('btncolor', {
									label: strings.BTNColorFieldLabel,
									selectedColor: this.properties.btncolor,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									disabled: false,
									// debounce: 1000,
									isHidden: false,
									alphaSliderHidden: false,
									style: PropertyFieldColorPickerStyle.Full,
									iconName: 'Precipitation',
									key: 'btncolorFieldId'
								}),
								PropertyPaneDropdown('linktxtcolor', {
									label: strings.LinkTxtColorFieldLabel,
									options: [
										{ key: "txtColorWhite", text: "White" },
										{ key: "txtColorBlack", text: "Black" }
									],
									selectedKey: "txtColorBlack"
								})
							]
						},
						{
							groupName: strings.AdditionalGroupName,
							isCollapsed: true,
							groupFields: [
								PropertyPaneToggle('showconfidential', {
									label: strings.ConfidentialFieldLabel,
									checked: true,
									onText: "Enable",
									offText: "Disable"
								}),
								PropertyPaneToggle('showannouncement', {
									label: strings.AnnouncementFieldLabel,
									checked: false,
									onText: "Enable",
									offText: "Disable"
								}),
								announcementMsgControl,
								announcementTypeControl,
								announceLinkToggleControl,
								announceLinkControl,
								announceLinkTextControl,
								PropertyPaneToggle('showhelplink', {
									label: strings.HelpLinkFieldLabel,
									checked: false,
									onText: "Enable",
									offText: "Disable"
								}),
								helplinkTextControl,
								helplinkURLControl,
								helpLinkInvertControl
							]
						}
					]
				}
			]
		};
	}
}


