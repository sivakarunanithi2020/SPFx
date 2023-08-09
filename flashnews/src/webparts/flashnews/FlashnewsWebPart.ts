import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FlashnewsWebPartStrings';
import Flashnews from './components/Flashnews';
import { IFlashnewsProps } from './components/IFlashnewsProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';  
import { IColumnReturnProperty,PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';  
import { sp } from '@pnp/sp';
import { SPService } from '../flashnews/service/service';   
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface IFlashnewsWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  FilterBy: string;
  condition: string;
  FilterValue:string;
  webUrl: string;
  Title:string;
  itemstore:[];

  lists: string;
  column: string;
  fields: string[];
  speed:number;
  direction:string;
  bgcolor:string;
  fgcolor:string;
  fontname:string;
  fontsize:string; 
  height:string;
  width:string;

  descbgcolor:string;
  descfgcolor:string;
  descfontsize:string;
  descfontname:string;
}

export default class FlashnewsWebPart extends BaseClientSideWebPart<IFlashnewsWebPartProps> {

  private _services: SPService = null;
  private _listFields: IPropertyPaneDropdownOption[] = [];

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      this._services = new SPService(this.context);
      this.getListFields = this.getListFields.bind(this);
    });
  }


  public render(): void {
    const element: React.ReactElement<IFlashnewsProps> = React.createElement(
      Flashnews,
      {
        description: this.properties.description,
        SiteUrl: this.properties.SiteUrl,
        ListName: this.properties.ListName,
        FilterBy: this.properties.FilterBy,
        condition: this.properties.condition,
        FilterValue:this.properties.FilterValue,
        webUrl: this.properties.webUrl,
        Title:this.properties.Title,

        context: this.context,
        list: this.properties.lists,
        column: this.properties.column,
        fields: this.properties.fields,
        speed: this.properties.speed,
        direction: this.properties.direction,
        bgcolor: this.properties.bgcolor,
        fgcolor:this.properties.fgcolor,
        fontname:this.properties.fontname,
        fontsize:this.properties.fontsize,
        height:this.properties.height,
        width:this.properties.width,

        descbgcolor:this.properties.descbgcolor,
        descfgcolor:this.properties.descfgcolor,
        descfontsize:this.properties.descfontsize,
        descfontname:this.properties.descfontname
      }
    );

    ReactDom.render(element, this.domElement);
  }


  public async getListFields() {
    if (this.properties.lists) {
      let allFields = await this._services.getFields(this.properties.lists);
      (this._listFields as []).length = 0;
      this._listFields.push(...allFields.map(field => ({ key: field.InternalName, text: field.Title })));
    }
  }

  private listConfigurationChanged(propertyPath: string, oldValue: any, newValue: any) {
    console.log("LIST FIELDS:", this._listFields);
    if (propertyPath === 'lists' && newValue) {
      this.properties.fields = [];
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /* private async  getColumnData(){
    // get all the items from a list
    const items: any[] = await sp.web.lists.getById(this.properties.lists).items();
    console.log(items);
  } */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
   // this.getListFields(); 
    //this.getColumnData();
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                // PropertyPaneTextField('SiteUrl', {
                //   label: "Site Url"
                // }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
              //    baseTemplate: 100,
                  onPropertyChange: this.listConfigurationChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  key: 'listPickerFieldId',
                }),
                // PropertyFieldMultiSelect('fields', {
                //   key: 'multiSelect',
                //   label: "Multi select list fields",
                //   options: this._listFields,
                //   selectedKeys: this.properties.fields
                // }),
                PropertyFieldColumnPicker('column', {
                  label: 'Select a column',
                  context: this.context as any,
                  selectedColumn: this.properties.column,
                  listId: this.properties.lists,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
              }),
              PropertyPaneSlider('speed',{  
                label:"Speed",  
                min:5,  
                max:100,  
                value:5,  
                showValue:true,  
                step:1                
              }),
              PropertyFieldColumnPicker('FilterBy', {
                label: 'Select Filter column',
                context: this.context as any,
                selectedColumn: this.properties.FilterBy,
                listId: this.properties.lists,
                disabled: false,
                orderBy: PropertyFieldColumnPickerOrderBy.Title,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'columnPickerFieldId',
                displayHiddenColumns: false,
                columnReturnProperty: IColumnReturnProperty["Internal Name"]
            }),
            PropertyPaneTextField('Title', {
                 label: "Title"
               }),
            PropertyPaneDropdown('direction', {
              label:'Direction',
              options:[
                {key:'left', text:'left'},
                {key:'right',text:'right'},
              ],
            }),
            // PropertyPaneDropdown('condition', {
            //   label: 'Condition',
            //   options: [
            //     { key: 'eq', text: 'equals' },
            //     { key: 'nq', text: 'not equal'},
            //     { key: 'ge', text: 'greater than or equal to' },
            //     { key: 'le', text: 'less than or equal to' }
            //   ],
            // }),
            // PropertyPaneTextField('FilterValue', {
            //   label: "Value"
            // }),
             ]
            }
          ]
        },  //page 1 ends here
        {
          // header: {
          //   description: "Design"
          // },
        
        groups: [
          {
            groupName: "Scrolling Text Design Configuration", 
            groupFields: [
              PropertyFieldColorPicker('bgcolor', {
                label: 'Background Color',
                selectedColor: this.properties.bgcolor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 1000,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldId'
              }),
              PropertyFieldColorPicker('fgcolor', {
                label: 'Text Color',
                selectedColor: this.properties.fgcolor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 1000,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldId'
              }),
              PropertyPaneTextField('fontname', {
                label: "Font Name"
              }),
              PropertyPaneTextField('fontsize', {
                label: "Font Size"
              }),
              PropertyPaneTextField('height', {
                label: "Height"
              }),
              PropertyPaneTextField('width', {
                label: "Width"
              }),
            ]}]
        }, // Page 2 end here
        {
          // header: {
          //   description: "Design"
          // },
        groups: [
          {
            groupName: "Description Text Design Configuration", 
            groupFields: [
              PropertyFieldColorPicker('descbgcolor', {
                label: 'Background Color',
                selectedColor: this.properties.descbgcolor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 1000,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldId'
              }),
              PropertyFieldColorPicker('descfgcolor', {
                label: 'Text Color',
                selectedColor: this.properties.descfgcolor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 1000,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldId'
              }),
              PropertyPaneTextField('descfontname', {
                label: "Font Name"
              }),
              PropertyPaneTextField('descfontsize', {
                label: "Font Size"
              })
            ]}]
        } // Page 3 end here
      ]
    };
  }
}
