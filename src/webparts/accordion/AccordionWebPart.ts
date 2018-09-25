import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import styles from './AccordionWebPart.module.scss';
import * as strings from 'AccordionWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';


require('jquery');
require('jqueryui');
import * as jQuery from 'jquery';
import 'jqueryui';


export interface IAccordionWebPartProps {
  description: string;
  tabs: any[];
  inline : boolean;
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {

  private guid: string;

  /**
   *
   */
  constructor(context?: IWebPartContext) {
    super();
    this.guid = this.getGuid();

    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);

    if (Environment.type !== EnvironmentType.ClassicSharePoint) {
      //Load the JQuery smoothness CSS file
      SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
    }
  }



  public render(): void {
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      var errorHtml = '';
      errorHtml += '<div style="color: red;">';
      errorHtml += '<div style="display:inline-block; vertical-align: middle;"><i class="ms-Icon ms-Icon--Error" style="font-size: 20px"></i></div>';
      errorHtml += '<div style="display:inline-block; vertical-align: middle;margin-left:7px;"><span>';
      errorHtml += strings.ErrorClassicSharePoint;
      errorHtml += '</span></div>';
      errorHtml += '</div>';
      this.domElement.innerHTML = errorHtml;
      return;
    }

    var html = '';

    //Define the main div
    html += '<div class="accordion" id="' + this.guid + '">';

    //Iterates on tabs

    if(!(this.properties.tabs && this.properties.tabs.length > 0)){
      return;
    }

    this.properties.tabs.map((tab: any, index: number) => {
      if (this.displayMode == DisplayMode.Edit) {
        //If diplay Mode is edit, include the textarea to edit the tab's content
        html += '<h3>' + (tab.Title != null ? tab.Title : '') + '</h3>';
        html += '<div style="min-height: 400px"><textarea name="' + this.guid + '-editor-' + index + '" id="' + this.guid + '-editor-' + index + '">' + (tab.Content != null ? tab.Content : '') + '</textarea></div>';
      }
      else {
        //Display Mode only, so display the tab content
        html += '<h3>' + (tab.Title != null ? tab.Title : '') + '</h3>';
        html += '<div>' + (tab.Content != null ? tab.Content : '') + '</div>';
      }
    });
    html += '</div>';

    //Flush the output HTML code
    this.domElement.innerHTML = html;

    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: false,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s'
      }
    };

    jQuery('#' + this.guid).accordion(accordionOptions);

    if (this.displayMode == DisplayMode.Edit) {
      //If the display mode is Edit, loads the CK Editor plugin
      var ckEditorCdn = '//cdn.ckeditor.com/4.6.2/full/ckeditor.js';
      //Loads the Javascript from the CKEditor CDN
      SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: 'CKEDITOR' }).then((CKEDITOR: any): void => {
        if (this.properties.inline == null || this.properties.inline === false) {
          //If mode is not inline, loads the script with the replace method
          for (var tab = 0; tab < this.properties.tabs.length; tab++) {
            CKEDITOR.replace(this.guid + '-editor-' + tab, {
              skin: 'moono-lisa,//cdn.ckeditor.com/4.6.2/full-all/skins/moono-lisa/'
            });
          }

        }
        else {
          //Mode is inline, so loads the script with the inline method
          for (var tab2 = 0; tab2 < this.properties.tabs.length; tab2++) {
            CKEDITOR.inline(this.guid + '-editor-' + tab2, {
              skin: 'moono-lisa,//cdn.ckeditor.com/4.6.2/full-all/skins/moono-lisa/'
            });
          }
        }
        //Catch the CKEditor instances change event to save the content
        for (var i in CKEDITOR.instances) {
          CKEDITOR.instances[i].on('change', (elm?, val?) => {
            //Updates the textarea
            elm.sender.updateElement();
            //Gets the value
            var value = ((document.getElementById(elm.sender.name)) as any).value;
            var id = elm.sender.name.split("-editor-")[1];
            //Save the content in properties
            this.properties.tabs[id].Content = value;
          });
        }
      });

    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  private s4(): string {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }

  private getGuid(): string {
    return this.s4() + this.s4() + '-' + this.s4() + '-' + this.s4() + '-' +
      this.s4() + '-' + this.s4() + this.s4() + this.s4();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldCollectionData('tabs', {
                  key: "tabsList",
                  label: "Accordion Data",
                  value: this.properties.tabs,
                  manageBtnLabel: "Manage Accordion Data",
                  panelHeader: "Create your Accordion Title",
                  fields: [
                    {
                      id: 'Title',
                      title: "Title",
                      type: CustomCollectionFieldType.string,
                      placeholder: "Enter your Accordion Title",
                      required: true
                    }
                  ]
                }),
                PropertyPaneToggle('inline', {
                  label: "Inline",
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
