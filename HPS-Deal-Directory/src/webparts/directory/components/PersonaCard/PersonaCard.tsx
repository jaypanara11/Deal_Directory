import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import {
  Log, Environment, EnvironmentType,
} from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  Persona,
  PersonaSize,
  DocumentCard,
  DocumentCardType,
  Icon,
} from 'office-ui-fabric-react';

const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string =
  '914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);
    console.log(props);
    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount() {
    if (Environment.type !== EnvironmentType.Local) {
      const sharedLibrary = await this._loadSPComponentById(
        LIVE_PERSONA_COMPONENT_ID
      );
      const livePersonaCard: any = sharedLibrary.LivePersonaCard;
      this.setState({ livePersonaCard: livePersonaCard });
    }

    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.css');
  }

  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */
  public componentDidUpdate(
    prevProps: IPersonaCardProps,
    prevState: IPersonaCardState
  ): void { }

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  private _LivePersonaCard() {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        // upn: this.props.profileProperties.Email,
        // onCardOpen: () => {
        //   console.log('LivePersonaCard Open');
        // },
        // onCardClose: () => {
        //   console.log('LivePersonaCard Close');
        // },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    var personaKey = Math.floor(Math.random() * 9999);
    return (
      <>
        <a href={this.props.tickerProperties.URL} data-interception="off" target="_blank">
          <DocumentCard
            className={styles.documentCard}
            type={DocumentCardType.normal}
          >
            <div className={styles.persona}>
              <Persona
                text={this.props.tickerProperties.DisplayType == "Company" ? this.props.tickerProperties.Issuer : this.props.tickerProperties.Title}
                size={PersonaSize.size72}
                imageShouldFadeIn={false}
                imageShouldStartVisible={false}
                key={personaKey}
              >
                <div>
                  {
                    this.props.tickerProperties.DisplayType == "Deal" && this.props.tickerProperties.SalesforceLink != "" && this.props.tickerProperties.SalesforceLink != undefined
                      ?
                      <a href={this.props.tickerProperties.SalesforceLink} target="_blank" data-interception="off">
                        Salesforce
                      </a>
                      :
                      <></>
                  }
                  {
                    this.props.tickerProperties.DisplayType == "Company" && this.props.tickerProperties.Sector != ""
                      ?
                      <div>
                        <b>Platform: </b>{this.props.tickerProperties.Sector}
                      </div>
                      :
                      <></>
                  }
                  {/* {
                    this.props.tickerProperties.Fund != ""
                      ?
                      <div>
                        <b>Fund: </b>{this.props.tickerProperties.Fund}
                      </div>
                      :
                      <></>
                  }

               
                {
                    this.props.tickerProperties.Country != ""
                      ?
                      <div>
                        <b>Country: </b>{this.props.tickerProperties.Country}
                      </div>
                      :
                      <></>
                  } */}

                </div>

                {/* const examplePersona: IPersonaSharedProps = {
    imageUrl: TestImages.personaFemale,
    imageInitials: 'AL',
    text: 'Annie Lindqvist',
    secondaryText: 'Software Engineer',
    tertiaryText: 'In a meeting',
    optionalText: 'Available at 4:00pm',
  }; */}
                {/* {this.props.profileProperties.WorkPhone ? (
          <div>
            <Icon iconName="Phone" style={{ fontSize: '12px' }} />
            <span style={{ marginLeft: 5, fontSize: '12px' }}>
              {' '}
              {this.props.profileProperties.WorkPhone}
            </span>
          </div>
        ) : (
            ''
          )}
        {this.props.profileProperties.Location ? (
          <div className={styles.textOverflow}>
            <Icon iconName="Poi" style={{ fontSize: '12px' }} />
            <span style={{ marginLeft: 5, fontSize: '12px' }}>
              {' '}
              {this.props.profileProperties.Location}
            </span>
          </div>
        ) : (
            ''
          )} */}
              </Persona>
            </div>
          </DocumentCard>
        </a>
      </>
    );
  }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Promise.reject(error);
      Log.error(EXP_SOURCE, error, this.props.context.serviceScope);
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    return (
      <div className={styles.personaContainer}>
        {this.state.livePersonaCard
          ? this._LivePersonaCard()
          : this._PersonaCard()}
      </div>
    );
  }
}
