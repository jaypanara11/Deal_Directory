import * as React from "react";
import styles from "./Directory.module.scss";
import { IDirectoryProps } from "./IDirectoryProps";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { spservices } from "../../../SPServices/spservices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings";
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox,
  Icon,
  Label,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
  Dropdown,
  IDropdownOption
} from "office-ui-fabric-react";

import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISPServices } from "../../../SPServices/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../../../SPServices/spMockServices";

const az: string[] = [
  "All",
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z"
];
const orderOptions: IDropdownOption[] = [
  { key: "FirstName", text: "First Name" },
  { key: "LastName", text: "Last Name" },
  { key: "Department", text: "Department" },
  { key: "Location", text: "Location" },
  { key: "JobTitle", text: "Job Title" }
];
export default class Directory extends React.Component<
  IDirectoryProps,
  IDirectoryState
> {
  private _services: ISPServices = null;

  constructor(props: IDirectoryProps) {
    super(props);

    this.state = {
      tickers: [],
      isLoading: true,
      errorMessage: "",
      hasError: false,
      indexSelectedKey: "All",
      searchString: "LastName",
      searchText: "",
      issuerSearchText: "",
      prjNameSearchText: "",
      issuerFilterOptions: [],
      prjNameFilterOptions: [],
      selectedDisplayOption: 'Issuer',
      selectedPlatformFilter: "",
      selectedFundFilter: "",
      selectedProjectNameFilter: "",
      selectedCrsPltfrmFilter: "",
      selectedDealTeamRgnFilter: "",
      selectedIndustryFilter: "",
      selectedCountryFilter: "",
      selectedStatus: "",
      selectedType: ""
    };
    if (Environment.type === EnvironmentType.Local) {
      this._services = new spMockServices();
    } else {
      this._services = new spservices(this.props.context);
    }
    // Register event handlers
    this._searchTickers = this._searchTickers.bind(this);
    this._selectedIndex = this._selectedIndex.bind(this);
    this._sortPeople = this._sortPeople.bind(this);
    this._searchBoxChanged = this._searchBoxChanged.bind(this);
  }

  /**
   *
   *
   * @memberof Directory
   */
  public async componentDidMount() {
    await this._searchTickers("All");
  }

  /**
   * Gets image base64
   * @param pictureUrl
   * @returns
   */
  private getImageBase64(pictureUrl: string): Promise<string> {
    return new Promise((resolve, reject) => {
      let image = new Image();
      image.addEventListener("load", () => {
        let tempCanvas = document.createElement("canvas");
        tempCanvas.width = image.width,
          tempCanvas.height = image.height,
          tempCanvas.getContext("2d").drawImage(image, 0, 0);
        let base64Str;
        try {
          base64Str = tempCanvas.toDataURL("image/png");
        } catch (e) {
          return "";
        }

        resolve(base64Str);
      });
      image.src = pictureUrl;
    });
  }
  private _searchBoxChanged(newvalue: string): void {
    this.setState({ searchText: newvalue }, () => this._searchTickers(newvalue));
  }

  private async _searchTickers(searchText: string) {
    searchText = searchText.trim().length > 0 ? searchText : "All";
    this.setState({
      isLoading: true,
      indexSelectedKey: searchText.substring(0, 1).toLocaleUpperCase(),
      searchString: "LastName"
    });

    try {
      const tickers = await this._services.searchTickers(
        searchText
      );
      if (tickers && tickers.length > 0) {
        for (let index = 0; index < tickers.PrimarySearchResults.length; index++) {
          let user: any = tickers.PrimarySearchResults[index];
          if (user.PictureURL) {
            user = { ...user, PictureURL: await this.getImageBase64(`/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}`) };
            tickers.PrimarySearchResults[index] = user;
          }
        }
      }

      this.setState({
        tickers: tickers,
        isLoading: false,
        errorMessage: "",
        hasError: false
      });
    } catch (error) {
      this.setState({ errorMessage: error.message, hasError: true });
    }
  }

  /**
   *
   *
   * @param {IDirectoryProps} prevProps
   * @param {IDirectoryState} prevState
   * @memberof Directory
   */
  public async componentDidUpdate(
    prevProps: IDirectoryProps,
    prevState: IDirectoryState
  ) {
    if (
      this.props.title != prevProps.title
    ) {
      await this._searchTickers("All");
    }
  }

  /**
   *
   *
   * @private
   * @param {string} sortField
   * @memberof Directory
   */
  private async _sortPeople(sortField: string) {
    let _tickers = this.state.tickers;
    _tickers = _tickers.sort((a: any, b: any) => {
      switch (sortField) {
        // Sorte by FirstName
        case "FirstName":
          const aFirstName = a.FirstName ? a.FirstName : "";
          const bFirstName = b.FirstName ? b.FirstName : "";
          if (aFirstName.toUpperCase() < bFirstName.toUpperCase()) {
            return -1;
          }
          if (aFirstName.toUpperCase() > bFirstName.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by LastName
        case "LastName":
          const aLastName = a.LastName ? a.LastName : "";
          const bLastName = b.LastName ? b.LastName : "";
          if (aLastName.toUpperCase() < bLastName.toUpperCase()) {
            return -1;
          }
          if (aLastName.toUpperCase() > bLastName.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by Location
        case "Location":
          const aBaseOfficeLocation = a.BaseOfficeLocation
            ? a.BaseOfficeLocation
            : "";
          const bBaseOfficeLocation = b.BaseOfficeLocation
            ? b.BaseOfficeLocation
            : "";
          if (
            aBaseOfficeLocation.toUpperCase() <
            bBaseOfficeLocation.toUpperCase()
          ) {
            return -1;
          }
          if (
            aBaseOfficeLocation.toUpperCase() >
            bBaseOfficeLocation.toUpperCase()
          ) {
            return 1;
          }
          return 0;
          break;
        // Sort by JobTitle
        case "JobTitle":
          const aJobTitle = a.JobTitle ? a.JobTitle : "";
          const bJobTitle = b.JobTitle ? b.JobTitle : "";
          if (aJobTitle.toUpperCase() < bJobTitle.toUpperCase()) {
            return -1;
          }
          if (aJobTitle.toUpperCase() > bJobTitle.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by Department
        case "Department":
          const aDepartment = a.Department ? a.Department : "";
          const bDepartment = b.Department ? b.Department : "";
          if (aDepartment.toUpperCase() < bDepartment.toUpperCase()) {
            return -1;
          }
          if (aDepartment.toUpperCase() > bDepartment.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        default:
          break;
      }
    });
    this.setState({ tickers: _tickers, searchString: sortField });
  }
  /**
   *
   *
   * @private
   * @param {PivotItem} [item]
   * @param {React.MouseEvent<HTMLElement>} [ev]
   * @memberof Directory
   */
  private _selectedIndex(item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) {
    this.setState({ searchText: "" }, () => this._searchTickers(item.props.itemKey));
  }
  /**
   *
   *
   * @returns {React.ReactElement<IDirectoryProps>}
   * @memberof Directory
   */
  public render(): React.ReactElement<IDirectoryProps> {
    const color = this.props.context.microsoftTeams ? "white" : "";

    const diretoryGrid =
      this.state.tickers && this.state.tickers.length > 0
        ? this.state.tickers.map((ticker: any) => {
          return (
            <PersonaCard
              context={this.props.context}
              tickerProperties={{
                Title: ticker.Title,
                URL: ticker.URL,
                Sector: ticker.Sector,
                Issuer: ticker.Issuer,
                ProjectName: ticker.ProjectName,
                Fund: ticker.Fund,
                CrossPlatformStatus: ticker.CrossPlatformStatus,
                DealTeamRegion: ticker.DealTeamRegion,
                Industry: ticker.Industry,
                Country: ticker.Country,
                SalesforceLink: ticker.SalesforceLink,
                DisplayType: ""
              }}
            />
          );
        })
        : [];

    return (
      <div className={styles.directory}>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
        />

        <div className={styles.searchBox}>
          <SearchBox
            placeholder={strings.SearchPlaceHolder}
            styles={{
              root: {
                minWidth: 180,
                maxWidth: 300,
                marginLeft: "auto",
                marginRight: "auto",
                marginBottom: 25
              }
            }}
            onSearch={this._searchTickers}
            onClear={() => {
              this._searchTickers("All");
            }}
            value={this.state.searchText}
            onChanged={this._searchBoxChanged}
          />
          <div>
            <Pivot
              styles={{
                root: {
                  paddingLeft: 10,
                  paddingRight: 10,
                  whiteSpace: "normal",
                  textAlign: "center"
                }
              }}
              linkFormat={PivotLinkFormat.tabs}
              selectedKey={this.state.indexSelectedKey}
              onLinkClick={this._selectedIndex}
              linkSize={PivotLinkSize.normal}
            >
              {az.map((index: string) => {
                return (
                  <PivotItem headerText={index} itemKey={index} key={index} />
                );
              })}
            </Pivot>
          </div>
        </div>
        {!this.state.tickers || this.state.tickers.length == 0 ? (
          <div className={styles.noUsers}>
            <Icon
              iconName={"ProfileSearch"}
              style={{ fontSize: "54px", color: color }}
            />
            <Label>
              <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                {strings.DirectoryMessage}
              </span>
            </Label>
          </div>
        ) : this.state.isLoading ? (
          <Spinner size={SpinnerSize.large} label={"searching ..."} />
        ) : this.state.hasError ? (
          <MessageBar messageBarType={MessageBarType.error}>
            {this.state.errorMessage}
          </MessageBar>
        ) : (
          <div className={styles.dropDownSortBy}>
            <Dropdown
              placeholder={""}
              label={""}
              options={orderOptions}
              selectedKey={this.state.searchString}
              onChange={(ev: any, value: IDropdownOption) => {
                this._sortPeople(value.key.toString());
              }}
              styles={{ dropdown: { width: 200 } }}
            />
            <div>{diretoryGrid}</div>
          </div>
        )}
      </div>
    );
  }
}
