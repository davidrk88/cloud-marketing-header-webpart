import * as React from 'react';
import styles from './CloudMarketingHeader.module.scss';
import { ICloudMarketingHeaderProps } from './ICloudMarketingHeaderProps';
import TargetAudience, {
  ITargetAudienceState
} from "../../../common/TargetAudience";
import { escape } from '@microsoft/sp-lodash-subset';

export interface ICloudMarketingAudienceTargetState extends ITargetAudienceState {
	confidentialBarToggle?: boolean;
}

export default class CloudMarketingHeader extends React.Component<ICloudMarketingHeaderProps, ICloudMarketingAudienceTargetState> {
	
	public constructor(props: ICloudMarketingHeaderProps) {
		super(props);

		this.state = { confidentialBarToggle: true };
	}

	private _getBGColorsClass(curColor) {
		switch(curColor) {
			case "bgColorWhite":
				return `${styles.bgColorWhite}`;
			case "bgColorBlack":
				return `${styles.bgColorBlack}`;
			case "bgColorRed":
				return `${styles.bgColorRed}`;
			case "bgColorOrange":
				return `${styles.bgColorOrange}`;
			case "bgColorYellow":
				return `${styles.bgColorYellow}`;
			case "bgColorGreen":
				return `${styles.bgColorGreen}`;
			case "bgColorBlue":
				return `${styles.bgColorBlue}`;
			case "bgColorPurple":
				return `${styles.bgColorPurple}`;
			case "bgColorPink":
				return `${styles.bgColorPink}`;
			case "bgColorCyan":
				return `${styles.bgColorCyan}`;
			default: //sets to site theme's dark color
				return `${styles.bgColorDefault}`;
		}
	}

	private _getTxtColorsClass(curColor) {
		switch(curColor) {
			case "txtColorBlack":
				return `${styles.txtColorBlack}`;
			case "txtColorRed":
				return `${styles.txtColorRed}`;
			case "txtColorOrange":
				return `${styles.txtColorOrange}`;
			case "txtColorYellow":
				return `${styles.txtColorYellow}`;
			case "txtColorGreen":
				return `${styles.txtColorGreen}`;
			case "txtColorBlue":
				return `${styles.txtColorBlue}`;
			case "txtColorPurple":
				return `${styles.txtColorPurple}`;
			case "txtColorPink":
				return `${styles.txtColorPink}`;
			case "txtColorCyan":
				return `${styles.txtColorCyan}`;
			default:
				return `${styles.txtColorWhite}`;
		}
	}

	private _getAnnouncementColorsClass(curColor) {
		switch(curColor) {
			case "colorInformational":
				return `${styles.colorInformational}`;
			case "colorWarning":
				return `${styles.colorWarning}`;
			case "colorCritical":
				return `${styles.colorCritical}`;
		}
	}

	public render(): React.ReactElement<ICloudMarketingHeaderProps> {
		// const renderTextCol = () => {}
		// console.log("Value from TEXT color dropdown: " + this.props.textcolor);
		// console.log("Value from BACKGROUND color dropdown: " + this.props.bgcolor);
		// console.log(this.props.headerLinksConfig);

		const confidentialBar = () => {
			return (
				<div className={ styles.confidentialBar }>
					<div>
						<div className={ styles.confidentialBarHeader }>
							<span className={ styles.confidentialBarIcon }>&nbsp;&nbsp;&#33;&nbsp;&nbsp;</span> <span>MICROSOFT CONFIDENTIAL - FOR INTERNAL USE ONLY</span>
						</div>
						<div>
							<span>This tool and its content are intended for Microsoft internal audience only.  Information contained herein should not be shared with anyone who does not have a business need to know.  To ensure compliance, please review the</span>
							&nbsp;<a href="https://microsoft.sharepoint.com/sites/CELAWeb-Compliance/SitePages/confidential-information.aspx" target="_blank" data-interception="off">confidential information policy</a>
						</div>
					</div>
					<div className={ styles.confidentialBarContentToggle } onClick={ () => this.setState({ confidentialBarToggle: false }) }>X</div>
				</div>
			);
		};

		const renderButtonLinks = () => {
			// let btnBGColor = this._getBGColorsClass(this.props.btncolor);
			let btnBGColor = '';
			let btntextColor = this._getTxtColorsClass(this.props.linktxtcolor);

			if (this.props.headerLinksConfig !== undefined) {
				// this.props.headerLinksConfig.sort((a, b) => (a.linkOrder > b.linkOrder) ? 1 : (a.linkOrder === b.linkOrder) ? ((a.linkText > b.linkText) ? 1 : -1) : -1 );
				return (
					<div id="wpCMHeader_btnlinks">
						{this.props.headerLinksConfig.filter(arrLinks => arrLinks.linkFlag).map((btnLink) =>
							<div className={ `${styles.btnContainer}` }>
								<TargetAudience pageContext={ this.props.pageContext } audienceTargets={ btnLink.audienceTargetsLinks }>
								<a key={btnLink.uniqueId} href={ btnLink.linkUrl } className={ `${styles.button} ${styles.btnLink} ${btnBGColor}` } style={{ backgroundColor: `${this.props.btncolor}` }} target="_blank" data-interception="off">
									<span className={ `${styles.label} ${btntextColor}` }>{ btnLink.linkText }</span>
								</a>
								</TargetAudience>
							</div>
						)}
					</div>
				);
			} else {
				return (
					<div></div>
				);
			}
		};

		const renderAnnouncementBar = () => {
			let announcementLink;
			let announcementType = this._getAnnouncementColorsClass(this.props.announcementtype);

			if(this.props.showannouncelink) {
				announcementLink = <a href={ this.props.announcementlink } target="_blank" data-interception="off">{ this.props.announcementlinktext }</a>;
			}

			return ( 
				<div className={ styles.announcementBarBG }>
					<div className={ `${styles.announcementBar} ${announcementType}` }>
						{escape(this.props.announcementmsg)} {announcementLink}
					</div>
				</div> 
			);
		};

		const renderHelpLink = () => {
			let helpLinkColor = (this.props.helplinkinvert ? `${styles.helpLinkInvert}` : '');
			let helpLinkPosition = (this.state.confidentialBarToggle ? `${styles.helpLinkPos1}` : `${styles.helpLinkPos2}`);

			return (
				<a href={ this.props.helplinkurl } target="_blank" data-interception="off" className={ `${styles.helpLink} ${helpLinkPosition} ${helpLinkColor}` }>
					{ this.props.helplinktext }
				</a>
			);
		};

		const renderMainContent = () => {
			// let txtColor = this._getTxtColorsClass(this.props.textcolor);
			let txtColor = '';
			return (
				<div>
					<div className={ styles.content }>
						<div className={ `${styles.title} ${styles.textLeft} ${txtColor}` } style={{ color: `${this.props.textcolor}` }}>{ this.props.title }</div>
						<p className={ `${styles.description} ${styles.textLeft} ${txtColor}` } style={{ color: `${this.props.textcolor}` }}>{ this.props.description }</p>
						<TargetAudience pageContext={ this.props.pageContext } audienceTargets={ this.props.audienceTargetsDesc }>
						<p className={ `${styles.description} ${styles.textLeft} ${txtColor}` } style={{ color: `${this.props.textcolor}` }}>{ this.props.descriptionat }</p>
						</TargetAudience>
						{ renderButtonLinks() }
					</div>
				</div>
			);
		};

		return (
			<div className={ styles.cloudMarketingHeader }>
				{(this.props.showconfidential && this.state.confidentialBarToggle ) ? confidentialBar() : null}
				<div id="cloudMarketingHeaderWebpartMain" className={ styles.imgFullWidth }>
					{ renderMainContent() }
					{ this.props.showhelplink ? renderHelpLink() : null }
				</div>
				{this.props.showannouncement ? renderAnnouncementBar() : null }
			</div>
		);
	}

}


