// import * as React from 'react';
// import styles from './CloudMarketingHeader.module.scss';
// import { ICloudMarketingHeaderProps } from './ICloudMarketingHeaderProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class SubMenuButtons extends React.Component {
  
//   public constructor(props) {
//     super(props);

//     this.state = { inputs: ['input-0'] };
//   }

//   public appendInput() {
//     var newInput = `input-${this.state.inputs.length}`;
//     this.setState(prevState => ({ inputs: prevState.inputs.concat([newInput]) }));
//   }


//   public render() {
//     return(
//       <div>
//         { 
//           this.state.inputs.map(input => <div className="submenuBtns"><input type="text" /></div>);
//         } 
//       </div>
//     );
//   }

// }