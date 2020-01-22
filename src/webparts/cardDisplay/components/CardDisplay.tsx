import * as React from 'react';
import styles from './CardDisplay.module.scss';
import { ICardDisplayProps } from './ICardDisplayProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {IProps} from '../../../classes/IProps'
import {IServiceClass} from '../../../classes/IService'
import {IStateCard} from '../../../classes/IState'
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardTitle,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardImage,
  IDocumentCardStyles,
  IDocumentCardActivityPerson,
  IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import {ImageProp} from '../../../common/Images/images'

export default class CardDisplay extends React.Component<IProps, IStateCard> {
  public ImageProp : ImageProp = new ImageProp();
  public serviceclass : IServiceClass = new IServiceClass();
  private options: any[] = [];
  public constructor(props: IProps, state: IStateCard){
    super(props);
    this.state = {
      fileschoice: [],
      sheetchoice: [],
      dpselectedItem: undefined,      
      dpsheetItem: undefined,
      dpselectedItems: [],
      dpsheetselectedItems: [],
      sheetname: "",
      filename: "",
      disabled: false,
      checked: false
    }

    this.serviceclass.getWelcomeMessageDetails(this.props.context, this.props.resturl+`&$top=`+this.props.itemcount).then((ticketitems: any) => {          
      console.log(ticketitems.value);
      console.log(this.props.context.pageContext.web.permissions);
      this.setState({
        fileschoice: ticketitems.value
      });
    });
  }

  public render(): React.ReactElement<IProps> {
  
    const people: IDocumentCardActivityPerson[] = [
      { name: 'Annie Lindqvist', profileImageSrc: this.ImageProp.personaFemale },
      { name: 'Roko Kolar', profileImageSrc: '', initials: 'RK' },
      { name: 'Aaron Reid', profileImageSrc: this.ImageProp.personaMale },
      { name: 'Christian Bergqvist', profileImageSrc: '', initials: 'CB' }
    ];
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewImageSrc: String(require('../../../common/Images/document-preview.png')),
          iconSrc: String(require('../../../common/Images/icon-ppt.png')),
          width: 318,
          height: 196,
          accentColor: '#ce4b1f'
        }
      ],
    };
    const cardStyles: IDocumentCardStyles = {
      root: { display: 'inline-block', marginRight: 20, marginBottom: 20, width: 320 }
    };
    return (          
      <div className="container">        
          {this.state.fileschoice.map((imageList) => {debugger;
            return (<DocumentCard styles={cardStyles} onClickHref="http://bing.com">
                      <DocumentCardImage height={150} imageFit={ImageFit.cover} imageSrc={this.props.assetfolderurl+`/document-preview2.png`} />
                      <DocumentCardDetails>
                        <DocumentCardTitle title={imageList.FileLeafRef} shouldTruncate />
                      </DocumentCardDetails>
                      <DocumentCardActivity activity="Modified March 13, 2018" people={people.slice(0, 3)} />
                    </DocumentCard>);
          })}
      </div>)                
  }
}
