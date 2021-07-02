import * as React from 'react';
import styles from './CloudMarketingHeader.module.scss';
import { ICloudMarketingHeaderProps } from './ICloudMarketingHeaderProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CloudMarketingHeader extends React.Component<ICloudMarketingHeaderProps, { confidentialBarToggle: boolean }> {

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
					<div className={ styles.confidentialBarHeader }>
						<span className={ styles.confidentialBarIcon }>&nbsp;&nbsp;&#33;&nbsp;&nbsp;</span> <span>MICROSOFT CONFIDENTIAL - FOR INTERNAL USE ONLY</span>
					</div>
					{ this.state.confidentialBarToggle 
						? 	<div className={ styles.confidentialBarContent }>
								<span>This tool and its content are intended for Microsoft internal audience only.  Information contained herein should not be shared with anyone who does not have a business need to know.  To ensure compliance, please review the</span> <a href="https://microsoft.sharepoint.com/sites/CELAWeb-Compliance/SitePages/confidential-information.aspx" target="_blank">confidential information policy</a>
							</div>
						:	null 
					}
					<div className={ styles.confidentialBarContentToggle }>
						<a href="#" onClick={ () => this.setState({ confidentialBarToggle: !this.state.confidentialBarToggle }) }>{ this.state.confidentialBarToggle ? `hide` : `show` }</a>
					</div>
				</div>
			);
		};

		const renderButtonLinks = () => {
			let btnBGColor = this._getBGColorsClass(this.props.btncolor);
			let btntextColor = this._getTxtColorsClass(this.props.linktxtcolor);

			if (this.props.headerLinksConfig !== undefined) {
				this.props.headerLinksConfig.sort((a, b) => (a.linkOrder > b.linkOrder) ? 1 : (a.linkOrder === b.linkOrder) ? ((a.linkText > b.linkText) ? 1 : -1) : -1 );
				return (
					<div>
						{this.props.headerLinksConfig.filter(arrLinks => arrLinks.linkFlag).map((btnLink) =>
							<a key={btnLink.uniqueId} href={ btnLink.linkUrl } className={ `${styles.button} ${btnBGColor}` } target="_blank">
								<span className={ `${styles.label} ${btntextColor}` }>{escape(btnLink.linkText)}</span>
							</a>
						)}
					</div>
				);
			} else {
				return (
					<div></div>
				);
			}
		};

		const renderMainContent = () => {
			let txtColor = this._getTxtColorsClass(this.props.textcolor);
			return (
				<div>
					<div className={ styles.content }>
						<div className={ `${styles.title} ${styles.textLeft} ${txtColor}` }>{escape(this.props.title)}</div>
						<p className={ `${styles.description} ${styles.textLeft} ${txtColor}` }>{escape(this.props.description)}</p>
						{ renderButtonLinks() }
					</div>
				</div>
			);
		};

		const renderAnnouncementBar = () => {
			let announcementBar;
			let announcementType = this._getAnnouncementColorsClass(this.props.announcementtype);

			if(this.props.showannouncement) {
				announcementBar = <div className={ `${styles.announcementBar} ${announcementType}` }>{escape(this.props.announcementmsg)}</div>;
			}

			return ( <div className={ styles.announcementBarBG }>{ announcementBar }</div> );
		};

		return (
			<div className={ styles.cloudMarketingHeader }>
				{ confidentialBar() }
				<div id="cloudMarketingHeaderWebpartMain" className={ styles.imgFullWidth }>
					{ renderMainContent() }
				</div>
				{ renderAnnouncementBar() }
			</div>
		);
	}

}


