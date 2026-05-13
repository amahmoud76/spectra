export interface IFilterState {
  searchText: string;
  documentType: string[];
  therapeuticArea: string[];
  subTherapeuticArea: string[];
  asset: string[];
  indication: string[];
  lineOfTherapy: string[];
  paid: string[];
  diseaseArea: string[];
  effectiveDateFrom: Date | null;
  effectiveDateTo: Date | null;
  uploadDateFrom: Date | null;
  uploadDateTo: Date | null;
}

// Default state — all filters empty
export const defaultFilterState: IFilterState = {
  searchText: "",
  documentType: [],
  therapeuticArea: [],
  subTherapeuticArea: [],
  asset: [],
  indication: [],
  lineOfTherapy: [],
  paid: [],
  diseaseArea: [],
  effectiveDateFrom: null,
  effectiveDateTo: null,
  uploadDateFrom: null,
  uploadDateTo: null,
};
