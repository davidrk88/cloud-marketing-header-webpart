import * as React from "react";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import spservices from "../service/spservices";
import { PageContext } from "@microsoft/sp-page-context";
export interface ITargetAudienceProps {
    pageContext:PageContext;
    // audienceTarget: IPropertyFieldGroupOrPerson[];
    // audienceTargets: string;
    audienceTargets: any[];
}
export interface ITargetAudienceState {
    canView?: boolean;
}
export default class TargetAudience extends React.Component<ITargetAudienceProps, ITargetAudienceState>{
    constructor(props: ITargetAudienceProps) {
        super(props);
        this.state = {
            canView: false
        } as ITargetAudienceState;

    }
    public componentDidMount(): void {
        //setting the state whether user has permission to view webpart
        this.checkUserCanViewWebpart();
    }
    public render(): JSX.Element {
        return (
            <div>
                {/*{this.props.audienceTarget ? */}
                {/*this.props.audienceTargets ? 
                    (this.state.canView ? this.props.children : ``)
                    :this.props.children
                */}
                {this.state.canView ? this.props.children : ``}
            </div>);
    }
    public checkUserCanViewWebpart(): void {
        const self = this;
        let proms = [];
        const errors = [];
        const _sv = new spservices();
        // if (self.props.audienceTarget) {
        if (self.props.audienceTargets) {
            if (self.props.audienceTargets.length > 0) {
                // self.props.audienceTarget.map((item) => {
                //     proms.push(_sv.isMember(item.fullName, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));
                // });
                // proms.push(_sv.isMember(self.props.audienceTargets, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));            
                // self.props.audienceTargets.split(',').map((item) => {
                self.props.audienceTargets.map((item) => {
                    console.log('Audience Target Item = ', item);
                    if (typeof item === 'string') {
                        console.log('Audience Target TYPE = ', typeof item);
                        proms.push(_sv.isMember(item, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));
                    } else {
                        proms.push(_sv.isMember(item.fullName, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));
                    }
                });
                console.log('Proms = ', proms);

                // Promise.race(
                //   proms.map(p => {
                //     return p.catch(err => {
                //       errors.push(err);
                //       if (errors.length >= proms.length) throw errors;
                //       return Promise.race(null);
                //     });
                //   })).then(val => {
                //     this.setState({ canView: true }); //atleast one promise resolved
                // });
                proms.map(p => {
                    p.then((value) => { // if promise resolves
                        this.setState({ canView: true });
                    });
                });
            } else {
                this.setState({ canView: true });    
            }
        } else {
            this.setState({ canView: true });
        }
    }
}