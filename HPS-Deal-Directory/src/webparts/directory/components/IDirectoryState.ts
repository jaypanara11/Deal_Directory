import {
  ComboBox,
  IComboBoxOption,
  SelectableOptionMenuItemType,
  IComboBoxStyles
} from 'office-ui-fabric-react';

export interface IDirectoryState {
  tickers: any;
  isLoading: boolean;
  errorMessage: string;
  hasError: boolean;
  indexSelectedKey: string;
  searchString: string;
  searchText: string;
  issuerSearchText: string;
  prjNameSearchText: string;
  issuerFilterOptions: IComboBoxOption[];
  prjNameFilterOptions: IComboBoxOption[];
  selectedDisplayOption: string;
  selectedPlatformFilter: string;
  selectedFundFilter: string;
  selectedProjectNameFilter: string;
  selectedCrsPltfrmFilter: string;
  selectedDealTeamRgnFilter: string;
  selectedIndustryFilter: string;
  selectedCountryFilter: string;
  selectedStatus: string;
  selectedType: string;
}
