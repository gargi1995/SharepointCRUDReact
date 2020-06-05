import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";  
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown'; 
export interface IGargiReactCrudState {  
  items: IDropdownOption[]; 
  Date: Date;
  user:string[];
  value?: { key: string | number | undefined };
  selectedTerms: IPickerTerms;
  managedmetadata:IPickerTerms;
}

