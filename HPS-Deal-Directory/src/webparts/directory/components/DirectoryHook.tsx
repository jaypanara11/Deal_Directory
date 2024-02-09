import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./Directory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { RefinersCard } from "./RefinersCard/RefinersCard";
import { spservices } from "../../../SPServices/spservices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings";
import {
    Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,
    Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, Link
} from "office-ui-fabric-react";
import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { debounce } from "throttle-debounce";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISPServices } from "../../../SPServices/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../../../SPServices/spMockServices";
import { IDirectoryProps } from './IDirectoryProps';
import Paging from './Pagination/Paging';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';

//import { initializeIcons } from '@fluentui/font-icons-mdl2';
import {
    ComboBox,
    IComboBoxOption,
    SelectableOptionMenuItemType,
    IComboBoxStyles} from 'office-ui-fabric-react';
import { IComboBoxProps } from 'office-ui-fabric-react';

import './style.css';

const refreshIcon: any = require('../../../Images/refreshicon.png');

const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const wrapStackTokens: IStackTokens = { childrenGap: 30 };
var tempIssuesdn = [], tempPrjNameDn = [], platformOptions = [], fundOptions = [], cpsOptions = [], dtrOptions = [], industryOptions = [], countryOptions = [];
const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
    let _services: ISPServices = null;
    if (Environment.type === EnvironmentType.Local) {
        _services = new spMockServices();
    } else {
        _services = new spservices(props.context);
    }
    const [az, setaz] = useState<string[]>([]);
    const [alphaKey, setalphaKey] = useState<string>('All');
    // const [departmentKey, setdepartmentKey] = useState<string>('');
    const [sectorKey, setsectorKey] = useState<string>('');
    const [issuerKey, setIssuerKey] = useState<string>('');
    const [prjnameKey, setPrjnameKey] = useState<string>('');
    const [fundKey, setFundeKey] = useState<string>('');
    const [crossPlatformStatusKey, setcrossPlatformStatusKey] = useState<string>('');
    const [dealTeamRgnnKey, setDealTeamRegionKey] = useState<string>('');
    const [industryKey, setIndustryKey] = useState<string>('');
    const [countryKey, setCountryKey] = useState<string>('');

    const [locationKey, setlocationKey] = useState<string>('');
    const [state, setstate] = useState<IDirectoryState>({
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
        selectedDisplayOption: "Company",
        selectedPlatformFilter: "",
        selectedFundFilter: "",
        selectedProjectNameFilter: "",
        selectedCrsPltfrmFilter: "",
        selectedDealTeamRgnFilter: "",
        selectedIndustryFilter: "",
        selectedCountryFilter: "",
        selectedStatus: "Active",
        selectedType: "Company"//07/02/24
    });
    const orderOptions: IDropdownOption[] = [
        { key: "FirstName", text: "First Name" },
        { key: "LastName", text: "Last Name" },
        { key: "Department", text: "Department" },
        { key: "Location", text: "Location" },
        { key: "JobTitle", text: "Job Title" }
    ];

    const itemStatusOptions: IDropdownOption[] = [
        { key: "Active", text: "Active" },
        { key: "Archive", text: "Archive" }
    ];

    const itemTypeOptions: IDropdownOption[] = [
        //07/02/2024
        { key: "Company", text: "Company" },
        { key: "Deal", text: "Deal" }
    ];
    const color = props.context.microsoftTeams ? "white" : "";
    // Paging
    const [pagedItems, setPagedItems] = useState<any[]>([]);
    const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 20);
    const [currentPage, setCurrentPage] = useState<number>(1);

    const _onPageUpdate = async (pageno?: number) => {
        var currentPge = (pageno) ? pageno : currentPage;
        var startItem = ((currentPge - 1) * pageSize);
        var endItem = currentPge * pageSize;
        let filItems = slice(state.tickers, startItem, endItem);
        setCurrentPage(currentPge);
        setPagedItems(filItems);
    };

    const displayOptions: IDropdownOption[] = [
        { key: "Company", text: "Company" },
        { key: "Project", text: "Project" }
    ];

    const tickerSectors = [], tickerIssuer = [], tickerProjectName = [], tickerFund = [], tickerCrossPlatform = [], tickerDealTeamRegion = [], tickerIndustry = [], tickerCountry = [];
    tempIssuesdn = [], tempPrjNameDn = [], platformOptions = [], fundOptions = [], cpsOptions = [], dtrOptions = [], industryOptions = [], countryOptions = [];
    //industryOptions.push({ key:"", text: "" });
    const diretoryGrid =
        pagedItems && pagedItems.length > 0
            ? pagedItems.map((ticker: any) => {
                if (ticker.Sector && tickerSectors.indexOf(ticker.Sector) < 0) {
                    tickerSectors.push(ticker.Sector);
                    platformOptions.push({ key: ticker.Sector, text: ticker.Sector });
                }
                if (ticker.Issuer && tickerIssuer.indexOf(ticker.Issuer) < 0) {
                    tickerIssuer.push(ticker.Issuer);
                    tempIssuesdn.push({ key: ticker.Issuer, text: ticker.Issuer });
                }
                if (ticker.ProjectName && tickerProjectName.indexOf(ticker.ProjectName) < 0) {
                    tickerProjectName.push(ticker.ProjectName);
                    tempPrjNameDn.push({ key: ticker.ProjectName, text: ticker.ProjectName });
                }
                if (ticker.Fund && tickerFund.indexOf(ticker.Fund) < 0) {
                    tickerFund.push(ticker.Fund);
                    fundOptions.push({ key: ticker.Fund, text: ticker.Fund });
                }
                if (ticker.CrossPlatformStatus && tickerCrossPlatform.indexOf(ticker.CrossPlatformStatus) < 0) {
                    tickerCrossPlatform.push(ticker.CrossPlatformStatus);
                    cpsOptions.push({ key: ticker.CrossPlatformStatus, text: ticker.CrossPlatformStatus });
                }
                if (ticker.DealTeamRegion && tickerDealTeamRegion.indexOf(ticker.DealTeamRegion) < 0) {
                    tickerDealTeamRegion.push(ticker.DealTeamRegion);
                    dtrOptions.push({ key: ticker.DealTeamRegion, text: ticker.DealTeamRegion });
                }
                if (ticker.Industry && tickerIndustry.indexOf(ticker.Industry) < 0) {
                    tickerIndustry.push(ticker.Industry);
                    industryOptions.push({ key: ticker.Industry, text: ticker.Industry });
                }
                if (ticker.Country && tickerCountry.indexOf(ticker.Country) < 0) {
                    tickerCountry.push(ticker.Country);
                    countryOptions.push({ key: ticker.Country, text: ticker.Country });
                }

                return (
                    <PersonaCard
                        context={props.context}
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
                            DisplayType: state.selectedType
                        }}
                    />
                );
            })
            : [];

    const refinerTickerSectorGrid =
        tickerSectors && tickerSectors.length > 0
            ? tickerSectors.sort().map((jobTitle: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setsectorKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, selectedItemText);
                        }}>{jobTitle}</Link>
                    </div>
                );
            })
            : [];

    const refinerIssuerSectorGrid =
        tickerIssuer && tickerIssuer.length > 0
            ? tickerIssuer.sort().map((issuer: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setIssuerKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", selectedItemText);
                        }}>{issuer}</Link>
                    </div>
                );
            })
            : [];
    // var tempIssuer = [];
    // tickerIssuer && tickerIssuer.length > 0
    //     ? tickerIssuer.sort().map((issuer: any) => {
    //         tempIssuer.push({ key: issuer, text: issuer });

    //     })
    //     : [];
    // setstate({ ...state, issuerFilterOptions: tempIssuer });


    const refinerProjectNameGrid =
        tickerProjectName && tickerProjectName.length > 0
            ? tickerProjectName.sort().map((PrjName: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", selectedItemText);
                        }}>{PrjName}</Link>
                    </div>
                );
            })
            : [];
    const refinerFundGrid =
        tickerFund && tickerFund.length > 0
            ? tickerFund.sort().map((fund: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", "", selectedItemText);
                        }}>{fund}</Link>
                    </div>
                );
            })
            : [];
    const refinerCrossPlatformGrid =
        tickerCrossPlatform && tickerCrossPlatform.length > 0
            ? tickerCrossPlatform.sort().map((CrossPlatform: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", "", "", selectedItemText);
                        }}>{CrossPlatform}</Link>
                    </div>
                );
            })
            : [];

    const dealTeamRegionGrid =
        tickerDealTeamRegion && tickerDealTeamRegion.length > 0
            ? tickerDealTeamRegion.sort().map((dealTeamRegion: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", "", "", "", selectedItemText);
                        }}>{dealTeamRegion}</Link>
                    </div>
                );
            })
            : [];

    const industryGrid =
        tickerIndustry && tickerIndustry.length > 0
            ? tickerIndustry.sort().map((indstry: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", "", "", "", "", selectedItemText);
                        }}>{indstry}</Link>
                    </div>
                );
            })
            : [];

    const countryGrid =
        tickerCountry && tickerCountry.length > 0
            ? tickerCountry.sort().map((cntry: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setPrjnameKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, "", "", "", "", "", "", "", selectedItemText);
                        }}>{cntry}</Link>
                    </div>
                );
            })
            : [];

    const _loadAlphabets = () => {
        let alphabets: string[] = [];
        alphabets.push("All");
        for (let i = 65; i < 91; i++) {
            alphabets.push(
                String.fromCharCode(i)
            );
        }
        setaz(alphabets);
    };


    const _searchByAlphabets = async (initialSearch: boolean, sectorparam?: string, issuerparam?: string, projectnameparam?: string, fundparam?: string, crossPlatformparam?: string, dealTeamRegionparam?: string, industryparam?: string, countryparam?: string) => {
        setstate({ ...state, isLoading: true, searchText: '', issuerSearchText: '', prjNameSearchText: '' });
        let tickers = null;
        if (initialSearch || (`${alphaKey}` == "All" && (!sectorKey && !sectorparam) && (!issuerKey && !issuerparam) && (!prjnameKey && !projectnameparam) && (!fundKey && !fundparam) && (!dealTeamRgnnKey && !dealTeamRegionparam) && (!crossPlatformStatusKey && !crossPlatformparam) && (!industryKey && !industryparam))) {
            tickers = await _services.searchTickersNew(state.selectedType, '', '', true);
        } else {
            let sectorSelectedKey = sectorKey ? sectorKey : (sectorparam ? sectorparam : '');
            let issuerSelectedKey = issuerKey ? issuerKey : (issuerparam ? issuerparam : '');
            let prjnameSelectedKey = prjnameKey ? prjnameKey : (projectnameparam ? projectnameparam : '');
            let fundSelectedKey = fundKey ? fundKey : (fundparam ? fundparam : '');
            let crossPlatformKey = crossPlatformStatusKey ? crossPlatformStatusKey : (crossPlatformparam ? crossPlatformparam : '');
            let dealTeamRegionKey = dealTeamRgnnKey ? dealTeamRgnnKey : (dealTeamRegionparam ? dealTeamRegionparam : '');
            let industryValKey = industryKey ? industryKey : (industryparam ? industryparam : '');
            let countryValKey = countryKey ? countryKey : (countryparam ? countryparam : '');

            var tempCntr = false;
            if (sectorparam != "" && sectorparam != undefined) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, sectorSelectedKey);
            }
            if (issuerparam) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", issuerSelectedKey);
            }
            if (projectnameparam) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", prjnameSelectedKey);
            }
            if (fundparam) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", "", fundSelectedKey);
            }
            if (crossPlatformKey) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", "", "", crossPlatformKey);
            }
            if (dealTeamRegionKey) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", "", "", "", dealTeamRegionKey);
            }

            if (industryValKey) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", "", "", "", "", industryValKey);
            }

            if (countryValKey) {
                tempCntr = true;
                tickers = await _services.searchTickersNew(state.selectedType, '', `${alphaKey}`, false, "", "", "", "", "", "", "", countryValKey);
            }

            if (!tempCntr) {
                tickers = await _services.searchTickersNew(state.selectedType, `${alphaKey}`, '', true);
            }
            //else 
        }
        setstate({
            ...state,
            searchText: '',
            issuerSearchText: '',
            prjNameSearchText: '',
            indexSelectedKey: initialSearch ? 'All' : (alphaKey ? alphaKey : state.indexSelectedKey),
            tickers: tickers,
            isLoading: false,
            errorMessage: "",
            hasError: false
        });
    };

    const _alphabetChange = async (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) => {
        setstate({ ...state, searchText: "", indexSelectedKey: item.props.itemKey, isLoading: true });
        setsectorKey('');
        setalphaKey(item.props.itemKey);
        if (item.props.itemKey == 'All') {
            // setdepartmentKey('');
            // setlocationKey('');
            setsectorKey('');
            _searchByAlphabets(true);
        }
        setCurrentPage(1);
        _searchByAlphabets(false);
    };

    let _searchTickers = async (searchText: string) => {
        try {
            setstate({ ...state, searchText: searchText, isLoading: true });
            if (searchText && searchText.length > 0) {

                const tickers = await _services.searchTickersNew(state.selectedType, '', searchText, false);
                setstate({
                    ...state,
                    searchText: searchText,
                    indexSelectedKey: 'All',
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
                setalphaKey('All');
            } else {
                setstate({ ...state, searchText: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _searchBoxChanged = (newvalue: string): void => {
        setCurrentPage(1);
        _searchTickers(newvalue);
        //setalphaKey("All");
        // setstate({ ...state, indexSelectedKey: "All" });
    };

    let _issuerSearchTickers = async (issuerSearchText: string) => {
        try {
            setstate({ ...state, searchText: issuerSearchText, isLoading: true });
            if (issuerSearchText && issuerSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, "", state.indexSelectedKey, false, state.selectedPlatformFilter, issuerSearchText);
                setstate({
                    ...state,
                    issuerSearchText: issuerSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
               // setalphaKey('All');
            } else {
                setstate({ ...state, issuerSearchText: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    // const _issuerSearchBoxChanged = (newvalue: string): void => {
    //     setCurrentPage(1);
    //     _issuerSearchTickers(newvalue);
    //     //setalphaKey("All");
    //     // setstate({ ...state, indexSelectedKey: "All" });
    // };

    const _issuerSearchBoxChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _issuerSearchTickers(option!.text);
        }
        else {
            _issuerSearchTickers("");
        }

    });// setSelectedKey(option!.key as string);

    const _platformDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _platformSearchTickers(option!.text);
        }
        else {
            _platformSearchTickers("");
        }

    });

    let _platformSearchTickers = async (platformSearchText: string) => {
        try {
            setstate({ ...state, searchText: platformSearchText, isLoading: true });
            if (platformSearchText && platformSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, platformSearchText, state.issuerSearchText);
                setstate({
                    ...state,
                    selectedPlatformFilter: platformSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
               // setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _fundDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _fundSearchTickers(option!.text);
        }
        else {
            _fundSearchTickers("");
        }

    });

    let _fundSearchTickers = async (fundSearchText: string) => {
        try {
            setstate({ ...state, searchText: fundSearchText, isLoading: true });
            if (fundSearchText && fundSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.selectedProjectNameFilter, state.issuerSearchText, fundSearchText);
                setstate({
                    ...state,
                    selectedFundFilter: fundSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
                //setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _cpsDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _cpsSearchTickers(option!.text);
        }
        else {
            _cpsSearchTickers("");
        }

    });

    let _cpsSearchTickers = async (cpsSearchText: string) => {
        try {
            setstate({ ...state, searchText: cpsSearchText, isLoading: true });
            if (cpsSearchText && cpsSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.issuerSearchText, state.selectedProjectNameFilter, state.selectedFundFilter, cpsSearchText);
                setstate({
                    ...state,
                    selectedCrsPltfrmFilter: cpsSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
               // setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _dtrDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _dtrSearchTickers(option!.text);
        }
        else {
            _dtrSearchTickers("");
        }

    });

    let _dtrSearchTickers = async (dtrSearchText: string) => {
        try {
            setstate({ ...state, searchText: dtrSearchText, isLoading: true });
            if (dtrSearchText && dtrSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.selectedProjectNameFilter, state.issuerSearchText, state.selectedFundFilter, state.selectedCrsPltfrmFilter, dtrSearchText);
                setstate({
                    ...state,
                    selectedDealTeamRgnFilter: dtrSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
              //  setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _industryDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _industrySearchTickers(option!.text);
        }
        else {
            _industrySearchTickers("");
        }

    });

    let _industrySearchTickers = async (indstrySearchText: string) => {
        try {
            setstate({ ...state, searchText: indstrySearchText, isLoading: true });
            if (indstrySearchText && indstrySearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.selectedProjectNameFilter, state.issuerSearchText, state.selectedFundFilter, state.selectedCrsPltfrmFilter, state.selectedDealTeamRgnFilter, indstrySearchText);
                setstate({
                    ...state,
                    selectedIndustryFilter: indstrySearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
              //  setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _countryDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _countrySearchTickers(option!.text);
        }
        else {
            _countrySearchTickers("");
        }

    });

    let _countrySearchTickers = async (countrySearchText: string) => {
        try {
            setstate({ ...state, searchText: countrySearchText, isLoading: true });
            if (countrySearchText && countrySearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.selectedProjectNameFilter, state.issuerSearchText, state.selectedFundFilter, state.selectedCrsPltfrmFilter, state.selectedDealTeamRgnFilter, state.selectedIndustryFilter, countrySearchText);
                setstate({
                    ...state,
                    selectedCountryFilter: countrySearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
                //setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _prjNameDropdownChanged: any = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _prjNameSearchTickers(option!.text);
        }
        else {
            _prjNameSearchTickers("");
        }

    });

    let _prjNameSearchTickers = async (projectNameSearchText: string) => {
        try {
            setstate({ ...state, searchText: projectNameSearchText, isLoading: true });
            if (projectNameSearchText && projectNameSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, state.selectedPlatformFilter, state.issuerSearchText, projectNameSearchText, state.selectedFundFilter, state.selectedCrsPltfrmFilter, state.selectedDealTeamRgnFilter, state.selectedIndustryFilter, state.selectedCountryFilter);
                setstate({
                    ...state,
                    selectedProjectNameFilter: projectNameSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
               // setalphaKey('All');
            } else {
                setstate({ ...state, selectedPlatformFilter: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _statusDropdownChanged: any = ((event, option) => {

        setstate({
            ...state,
            selectedStatus: option!.text
        });

    });

    const _typeDropdownChanged: any = (async (event, option) => {

        // setstate({
        //     ...state,
        //     selectedType: option!.text
        // });

        setCurrentPage(1);

        // _projectSearchTickers("");

        const tickers = await _services.searchTickersNew(option!.text, '', "", false);
        setstate({
            ...state,
            selectedPlatformFilter: "",
            selectedFundFilter: "",
            selectedProjectNameFilter: "",
            selectedCrsPltfrmFilter: "",
            selectedDealTeamRgnFilter: "",
            selectedIndustryFilter: "",
            selectedCountryFilter: "",
            indexSelectedKey: 'All',
            tickers: tickers,
            isLoading: false,
            errorMessage: "",
            hasError: false,
            selectedType: option!.text
        });
        setalphaKey('All');

    });

    const _issuerSearchReset: any = ((event, option) => {
        setCurrentPage(1);

        _issuerSearchTickers("");

    });

    let _projectSearchTickers = async (projectSearchText: string) => {
        try {
            setstate({ ...state, searchText: projectSearchText, isLoading: true });
            if (projectSearchText && projectSearchText.length > 0) {
                const tickers = await _services.searchTickersNew(state.selectedType, '', state.indexSelectedKey, false, "", "", projectSearchText);
                setstate({
                    ...state,
                    prjNameSearchText: projectSearchText,
                    indexSelectedKey: state.indexSelectedKey,
                    tickers: tickers,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
               // setalphaKey('All');
            } else {
                setstate({ ...state, prjNameSearchText: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    // const _projectSearchBoxChanged = (newvalue: string): void => {
    //     setCurrentPage(1);
    //     _projectSearchTickers(newvalue);
    //     //setalphaKey("All");
    //     // setstate({ ...state, indexSelectedKey: "All" });
    // };
    const _projectSearchBoxChanged: IComboBoxProps['onChange'] = ((event, option) => {
        setCurrentPage(1);
        if (option != undefined) {
            _projectSearchTickers(option!.text);
        }
        else {
            _projectSearchTickers("");
        }

    });

    const _prjNameSearchReset: any = ((event, option) => {
        setCurrentPage(1);

        _projectSearchTickers("");

    });

    const resetSearch: any = (async (event, option) => {
        setCurrentPage(1);

        // _projectSearchTickers("");

        const tickers = await _services.searchTickersNew("Company", '', "", false);
        setstate({
            ...state,
            selectedPlatformFilter: "",
            selectedFundFilter: "",
            selectedProjectNameFilter: "",
            selectedCrsPltfrmFilter: "",
            selectedDealTeamRgnFilter: "",
            selectedIndustryFilter: "",
            selectedCountryFilter: "",
            indexSelectedKey: 'All',
            tickers: tickers,
            isLoading: false,
            errorMessage: "",
            hasError: false,
            selectedType: "Company",
            searchText: "",
            searchString: "",
            issuerSearchText: ""
        });
        setalphaKey('All');

    });

    _searchTickers = debounce(500, _searchTickers);
    _issuerSearchTickers = debounce(500, _issuerSearchTickers);
    _projectSearchTickers = debounce(500, _projectSearchTickers);

    const _sortPeople = async (sortField: string) => {
        let _tickers = [...state.tickers];
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
        setstate({ ...state, tickers: _tickers, searchString: sortField });
    };

    // const _DisplayOptionChanged: IDropdownOption['onChange'] = ((event, option) => {
    //     setCurrentPage(1);

    //     setstate({ ...state, selectedDisplayOption: option!.text });

    // });

    useEffect(() => {
        setPageSize(props.pageSize);
        if (state.tickers) _onPageUpdate();
    }, [state.tickers, props.pageSize]);

    useEffect(() => {
        if (((alphaKey.length > 0 && alphaKey != "0") || (!sectorKey)) && !state.searchText) _searchByAlphabets(false);
    }, [alphaKey]);

    useEffect(() => {
        _loadAlphabets();
        _searchByAlphabets(true);
    }, [props]);

    return (
        <div className={styles.directory}>
            <WebPartTitle displayMode={props.displayMode} title={props.title}
                updateProperty={props.updateProperty} />

            <div className={styles.deal_dire_maindiv}>
                <div className={styles.deal_dire_filterpanelOne}>
                    {/* <div className={styles.deal_dire_itemtype}>
                        <Dropdown
                            options={itemStatusOptions}
                            selectedKey={state.selectedStatus}
                            onChange={_statusDropdownChanged}
                        />
                    </div> */}
                    <div className={styles.deal_dire_displaytype}>
                        <Dropdown
                            options={itemTypeOptions}
                            selectedKey={state.selectedType}
                            onChange={_typeDropdownChanged}
                        />
                    </div>
                    <div className={styles.deal_dire_mainsearch}>
                        {/* <input type="text" placeholder="Search for issuer" className={styles.mainsearch_txt}></input> */}
                        {
                            state.selectedType == "Company"  // 07/02/24
                                ?
                                <SearchBox placeholder="Search for Company"  // 07/02/24
                                    onSearch={_searchTickers}
                                    value={state.searchText}
                                    onChanged={_searchBoxChanged}
                                    className={styles.mainsearch_txt}
                                />
                                :
                                <SearchBox placeholder="Search for Deal"
                                    onSearch={_searchTickers}
                                    value={state.searchText}
                                    onChanged={_searchBoxChanged}
                                    className={styles.mainsearch_txt}
                                />
                        }

                    </div>
                    <div className={styles.deal_dire_resetsearch}>
                        {/* <input type="button" value="Reset" className={styles.reset_btn}></input> */}
                        <img src={refreshIcon} className={styles.reset_btn} onClick={resetSearch}></img>
                    </div>
                </div>
                <div className={styles.deal_dire_filterpanelTwo}>
                    <Pivot className={styles.alphabets} linkFormat={PivotLinkFormat.tabs}
                        selectedKey={state.indexSelectedKey} onLinkClick={_alphabetChange}
                        linkSize={PivotLinkSize.normal} >
                        {az.map((index: string) => {
                            return (
                                <PivotItem headerText={index} itemKey={index} key={index} />
                            );
                        })}
                    </Pivot>
                </div>
                <div className={styles.deal_dire_filterpanelThree}>
                    <div className={styles.deal_dire_filterIssuer}>
                        {/* <ComboBox
                            selectedKey={state.issuerSearchText}
                            allowFreeform
                            autoComplete="on"
                            options={tempIssuesdn}
                            placeholder="Select Issuer"
                            onChange={_issuerSearchBoxChanged}
                            className={styles.deal_dire_filterDropdown}                           
                        /> */}
                        <Dropdown
                            options={tempIssuesdn}
                            placeholder="Select Company"
                            onChange={_issuerSearchBoxChanged}
                            selectedKey={state.issuerSearchText}
                        />
                    </div>
                    <div className={styles.deal_dire_filterPlatform}>
                        <Dropdown
                            options={platformOptions}
                            placeholder="Select Platform"
                            selectedKey={state.selectedPlatformFilter}
                            onChange={_platformDropdownChanged}
                        />
                    </div>
                    {/* <div className={styles.deal_dire_filterFund}>
                        <Dropdown
                            options={fundOptions}
                            placeholder="Select Fund"
                            selectedKey={state.selectedFundFilter}
                            onChange={_fundDropdownChanged}
                        />
                    </div> */}
                    <div className={styles.deal_dire_filterCPS}>
                        <Dropdown
                            options={cpsOptions}
                            placeholder="Select Cross Platform Status"
                            selectedKey={state.selectedCrsPltfrmFilter}
                            onChange={_cpsDropdownChanged}
                        />
                    </div>
                    {/* <div className={styles.deal_dire_filterDTR}>
                        <Dropdown
                            options={dtrOptions}
                            placeholder="Select Deal Team Region"
                            selectedKey={state.selectedDealTeamRgnFilter}
                            onChange={_dtrDropdownChanged}
                        />
                    </div> */}
                    <div className={styles.deal_dire_filterIndustry}>
                        <Dropdown
                            options={industryOptions}
                            placeholder="Select Industry"
                            selectedKey={state.selectedIndustryFilter}
                            onChange={_industryDropdownChanged}
                        />
                    </div>
                    {/* <div className={styles.deal_dire_filterCountry}>
                        <Dropdown
                            options={countryOptions}
                            placeholder="Select Country"
                            selectedKey={state.selectedCountryFilter}
                            onChange={_countryDropdownChanged}
                        />
                    </div> */}
                    <div className={styles.deal_dire_filterPrjName}>

                        <Dropdown
                            options={tempPrjNameDn}
                            placeholder="Select Project Name"
                            selectedKey={state.selectedProjectNameFilter}
                            onChange={_prjNameDropdownChanged}
                        />
                    </div>
                </div>
            </div>


            {/* <div className={styles.searchBox}> */}
            {/* <SearchBox placeholder="Search for Issuer" className={styles.searchTextBox}
                    onSearch={_searchTickers}
                    value={state.searchText}
                    onChange={_searchBoxChanged} /> */}
            {/* <div> */}
            {/* <Dropdown
                        options={displayOptions}
                        selectedKey={state.selectedDisplayOption}
                        id="selectDisplayOptions"
                        onChange={_DisplayOptionChanged}
                    ></Dropdown> */}
            {/* <Pivot className={styles.alphabets} linkFormat={PivotLinkFormat.tabs}
                        selectedKey={state.indexSelectedKey} onLinkClick={_alphabetChange}
                        linkSize={PivotLinkSize.normal} >
                        {az.map((index: string) => {
                            return (
                                <PivotItem headerText={index} itemKey={index} key={index} />
                            );
                        })}
                    </Pivot> */}
            {/* </div>
            </div> */}
            {state.isLoading ? (
                <div style={{ marginTop: '10px' }}>
                    <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
                </div>
            ) : (
                <>
                    {state.hasError ? (
                        <div style={{ marginTop: '10px' }}>
                            <MessageBar messageBarType={MessageBarType.error}>
                                {state.errorMessage}
                            </MessageBar>
                        </div>
                    ) : (
                        <>
                            {!pagedItems || pagedItems.length == 0 ? (
                                <div className={styles.noUsers}>
                                    <Icon
                                        iconName={"ProfileSearch"}
                                        style={{ fontSize: "54px", color: color }}
                                    />
                                    <Label>
                                        <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                                            {
                                                state.selectedType == "Deal"
                                                ?
                                                 "No deal found in directory"
                                                :
                                                    "No issuer found in directory"
                                            }
                                            {/* {strings.DirectoryMessage} */}
                                        </span>
                                    </Label>
                                </div>
                            ) : (
                                <>
                                    <div style={{ width: '100%', display: 'inline-block' }}>

                                        <div style={{ width: '100%', display: 'inline-block' }}>
                                            <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                                <div>
                                                    {diretoryGrid}
                                                </div>
                                            </Stack>
                                        </div>
                                    </div>
                                    <div style={{ width: '100%', display: 'inline-block' }}>
                                        <Paging
                                            totalItems={state.tickers ? state.tickers.length : 0}
                                            itemsCountPerPage={pageSize}
                                            onPageUpdate={_onPageUpdate}
                                            currentPage={currentPage} />
                                    </div>
                                </>
                            )}
                        </>
                    )}
                </>
            )}
        </div>
    );
};

export default DirectoryHook;