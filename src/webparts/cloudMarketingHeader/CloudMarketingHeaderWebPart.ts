import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
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
import { 
	FieldCollectionData, 
	CustomCollectionFieldType as FieldCollectionCustomType 
} from '@pnp/spfx-controls-react/lib/FieldCollectionData';
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
	titlecolor: string;
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
	showcustominfo: boolean;
	custominfocontent: string;
	custominfobgcolor: string;
	audienceTargetsInfo: IPropertyFieldGroupOrPerson[];
	custominfohandler: any;
	showcustominfo2: boolean;
	custominfo2content: string;
	custominfo2bgcolor: string;
	audienceTargetsInfo2: IPropertyFieldGroupOrPerson[];
	custominfo2handler: any;
	context: any;
	headerLinksConfig: any[];
	audienceTargetsDesc: IPropertyFieldGroupOrPerson[];
	audienceTargetsDescAT: IPropertyFieldGroupOrPerson[];
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
		let pageDisplayMode = this.displayMode == DisplayMode.Edit ? true : false;

		const element: React.ReactElement<ICloudMarketingHeaderProps> = React.createElement(
			CloudMarketingHeader,
			{
				filePickerResult: this.properties.filePickerResult,
				gradientcolor: this.properties.gradientcolor,
				title: this.properties.title,
				description: this.properties.description,
				descriptionat: this.properties.descriptionAT,
				titlecolor: this.properties.titlecolor,
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
				showcustominfo: this.properties.showcustominfo,
				custominfocontent: this.properties.custominfocontent,
				custominfobgcolor: this.properties.custominfobgcolor,
				audienceTargetsInfo: this.properties.audienceTargetsInfo,
				custominfohandler: this.onCustomInfoChange,
				showcustominfo2: this.properties.showcustominfo2,
				custominfo2content: this.properties.custominfo2content,
				custominfo2bgcolor: this.properties.custominfo2bgcolor,
				audienceTargetsInfo2: this.properties.audienceTargetsInfo2,
				custominfo2handler: this.onCustomInfo2Change,
				context: this.context,
				pagedisplaymode: pageDisplayMode,
				headerLinksConfig: this.properties.headerLinksConfig,
				audienceTargetsDesc: this.properties.audienceTargetsDesc,
				audienceTargetsDescAT: this.properties.audienceTargetsDescAT,
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
		this.domElement.querySelector('#customInfoPanel').setAttribute("style", "background-color: " + this.properties.custominfobgcolor + ";");
		this.domElement.querySelector('#customInfo2Panel').setAttribute("style", "background-color: " + this.properties.custominfo2bgcolor + ";");
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	private onCustomInfoChange = (newText: string) => {
		this.properties.custominfocontent = newText;
		return newText;
	}

	private onCustomInfo2Change = (newText: string) => {
		this.properties.custominfo2content = newText;
		return newText;
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		// Dynamic Configuration Pane controls
		let announcementMsgControl: any = [];
		let announcementTypeControl: any = [];
		let announceLinkToggleControl: any = [];
		let announceLinkControl: any = [];
		let announceLinkTextControl: any = [];
		let helplinkTextControl: any = [];
		let helplinkURLControl: any = [];
		let helpLinkInvertControl: any = [];
		let customInfoBGColorControl: any = [];
		let customInfoAudienceControl: any = [];
		let customInfo2BGColorControl: any = [];
		let customInfo2AudienceControl: any = [];

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

		if(this.properties.showcustominfo) {
			customInfoBGColorControl = PropertyFieldColorPicker('custominfobgcolor', {
				label: strings.CustomInfoBGColorLabel,
				selectedColor: this.properties.custominfobgcolor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				disabled: false,
				// debounce: 1000,
				isHidden: false,
				alphaSliderHidden: false,
				style: PropertyFieldColorPickerStyle.Inline,
				iconName: 'Color',
				key: 'infocolorFieldId'
			});

			customInfoAudienceControl = PropertyFieldPeoplePicker('audienceTargetsInfo', {
				label: strings.CustomInfoTargetPickerLabel,
				initialData: this.properties.audienceTargetsInfo,
				allowDuplicate: false,
				principalType: [PrincipalTypeProp.SharePoint],
				onPropertyChange: this.onPropertyPaneFieldChanged,
				context: this.context,
				properties: this.properties,
				onGetErrorMessage: null,
				deferredValidationTime: 0,
				key: 'infopeopleFieldId'
			});
		}

		if(this.properties.showcustominfo2) {
			customInfo2BGColorControl = PropertyFieldColorPicker('custominfo2bgcolor', {
				label: strings.CustomInfoBGColorLabel,
				selectedColor: this.properties.custominfo2bgcolor,
				onPropertyChange: this.onPropertyPaneFieldChanged,
				properties: this.properties,
				disabled: false,
				// debounce: 1000,
				isHidden: false,
				alphaSliderHidden: false,
				style: PropertyFieldColorPickerStyle.Inline,
				iconName: 'Color',
				key: 'info2colorFieldId'
			});

			customInfo2AudienceControl = PropertyFieldPeoplePicker('audienceTargetsInfo2', {
				label: strings.CustomInfoTargetPickerLabel,
				initialData: this.properties.audienceTargetsInfo2,
				allowDuplicate: false,
				principalType: [PrincipalTypeProp.SharePoint],
				onPropertyChange: this.onPropertyPaneFieldChanged,
				context: this.context,
				properties: this.properties,
				onGetErrorMessage: null,
				deferredValidationTime: 0,
				key: 'info2peopleFieldId'
			});
		}

		return {
			pages: [
				{
					displayGroupsAsAccordion: false,
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							isCollapsed: false,
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
										{ key: "bgColorBlack", text: "Black" },
										{ key: "bgColorWhite", text: "White" }
									],
									selectedKey: "bgColorBlack"
								}),
								PropertyPaneTextField('title', {
									label: strings.TitleFieldLabel
								}),
								PropertyFieldColorPicker('titlecolor', {
									label: strings.TitleColorFieldLabel,
									selectedColor: this.properties.titlecolor,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									disabled: false,
									// debounce: 1000,
									isHidden: false,
									alphaSliderHidden: false,
									style: PropertyFieldColorPickerStyle.Inline,
									iconName: 'Color',
									key: 'titlecolorFieldId'
								}),
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
									multiline: true
								}),
								PropertyFieldPeoplePicker('audienceTargetsDesc', {
									label: strings.DescriptionPickerLabel,
									initialData: this.properties.audienceTargetsDesc,
									allowDuplicate: false,
									principalType: [PrincipalTypeProp.SharePoint],
									onPropertyChange: this.onPropertyPaneFieldChanged,
									context: this.context,
									properties: this.properties,
									onGetErrorMessage: null,
									deferredValidationTime: 0,
									key: 'peopleDescFieldId'
								}),
								PropertyPaneTextField('descriptionAT', {
									label: strings.DescriptionATFieldLabel,
									multiline: true
								}),
								PropertyFieldPeoplePicker('audienceTargetsDescAT', {
									label: strings.DescriptionATPickerLabel,
									initialData: this.properties.audienceTargetsDescAT,
									allowDuplicate: false,
									principalType: [PrincipalTypeProp.SharePoint],
									onPropertyChange: this.onPropertyPaneFieldChanged,
									context: this.context,
									properties: this.properties,
									onGetErrorMessage: null,
									deferredValidationTime: 0,
									key: 'peopleFieldId'
								}),
								PropertyFieldColorPicker('textcolor', {
									label: strings.TextColorFieldLabel,
									selectedColor: this.properties.textcolor,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									disabled: false,
									// debounce: 1000,
									isHidden: false,
									alphaSliderHidden: false,
									style: PropertyFieldColorPickerStyle.Inline,
									iconName: 'Color',
									key: 'textcolorFieldId'
								})
							]
						}
					]
				},
				{
					displayGroupsAsAccordion: false,
					header: {
						description: strings.LinksPaneDescription
					},
					groups: [
						{
							groupName: strings.LinksGroupName,
							isCollapsed: false,
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
															let tempTargets = [];
															items.map((target) => {
																tempTargets.push(target.secondaryText);
															});
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
											id: "submenuBtns",
											title: "Sub-menu Buttons",
											type: CustomCollectionFieldType.custom,
											onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
												return (  
													React.createElement(FieldCollectionData, {
														key: itemId,
														// label: "Test Submenu Collect",
														manageBtnLabel: "Configure",
														// panelHeader: "Submenu Buttons",
														panelHeader: item.linkText+' sub-menu link settings',
														enableSorting: true,
														value: item.submenuBtns,
														onChanged: (values) => { 
															console.log(values);
															let tempValues = [];
															values.map((btnLink) => {
																tempValues.push(btnLink);
															});
															item.submenuBtns = tempValues;
															onUpdate(field.id, item.submenuBtns);
														},
														fields: [
															{id: "sublinktext", title: "Link Text", type: FieldCollectionCustomType.string, required: true},
															{id: "sublinkurl", title: "URL", type: FieldCollectionCustomType.string}
														]
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
										{
											id: "linkFlag",
											title: "Enable",
											defaultValue: true,
											type: CustomCollectionFieldType.boolean
										}
									],
									disabled: false
								}),
								PropertyFieldColorPicker('btncolor', {
									label: strings.BTNColorFieldLabel,
									selectedColor: this.properties.btncolor,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									disabled: false,
									// debounce: 1000,
									isHidden: false,
									alphaSliderHidden: false,
									style: PropertyFieldColorPickerStyle.Inline,
									iconName: 'Color',
									key: 'btncolorFieldId'
								}),
								PropertyPaneDropdown('linktxtcolor', {
									label: strings.LinkTxtColorFieldLabel,
									options: [
										{ key: "txtColorBlack", text: "Black" },
										{ key: "txtColorWhite", text: "White" }
									],
									selectedKey: "txtColorBlack"
								})
							]
						}
					]
				},
				{
					displayGroupsAsAccordion: true,
					header: {
						description: strings.AdditionalInfoDesc
					},
					groups: [
						{
							groupName: strings.NotificationsGroupName,
							isCollapsed: false,
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
						},
						{
							groupName: strings.CustomInfoGroupName,
							isCollapsed: false,
							groupFields: [
								PropertyPaneToggle('showcustominfo', {
									label: strings.CustomInfoFieldLabel,
									checked: false,
									onText: "Enable",
									offText: "Disable"
								}),
								customInfoBGColorControl,
								customInfoAudienceControl
							]
						},
						{
							groupName: strings.CustomInfo2GroupName,
							isCollapsed: false,
							groupFields: [
								PropertyPaneToggle('showcustominfo2', {
									label: strings.CustomInfoFieldLabel,
									checked: false,
									onText: "Enable",
									offText: "Disable"
								}),
								customInfo2BGColorControl,
								customInfo2AudienceControl
							]
						}
					]
				}
			]
		};
	}
}


