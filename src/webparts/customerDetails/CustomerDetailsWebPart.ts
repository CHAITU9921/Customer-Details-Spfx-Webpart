import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustomerDetailsWebPart.module.scss';
import * as strings from 'CustomerDetailsWebPartStrings';
import { Web }  from 'sp-pnp-js';

export interface ICustomerDetailsWebPartProps {
  description: string;
}

export default class CustomerDetailsWebPart extends BaseClientSideWebPart<ICustomerDetailsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.customerDetails} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
    <div style="margin: 0; font-family: 'Arial', sans-serif; background: linear-gradient(to right, #1e90ff, #5A7DBD);  overflow-x: hidden;">
    <div style="padding: 20px; text-align: ">
   
      <section class="customerDetails_56633758 ">
    
       <div style="font-family: 'Arial', sans-serif;  margin: 0; padding: 0; display: flex; align-items: center; justify-content: center; height: 1000px;">
       <div action="/submit_form" style="background-color: rgba(200,200,220,0.2); padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); width: 450px;height: 930px;">
       <center><h1 style="color: #020202; font-size: 30px; margin: 0px;">Customer Entry Form</h1></center>
       <label for="companyName" style="display: block; margin-bottom: 8px; color: #020202;">Company Name <span style="color: red;">*</span> :</label>
           <input type="text" id="companyName" name="companyName" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
   
           <label for="customerName" style="display: block; margin-bottom: 8px; color: #020202;">Customer Name / Person Name <span style="color: red;">*</span> :</label>
           <input type="text" id="customerName" name="customerName" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
   
           <label for="address" style="display: block; margin-bottom: 8px; color: #020202;">Address :</label>
           <textarea id="address" name="address" rows="4" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;"></textarea>
   
           <label for="City" style="display: block; margin-bottom: 8px; color: #020202;">City :</label>
           <input id="City" type="text" name="City" rows="4" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;"></input>
          
           <label for="State" style="display: block; margin-bottom: 8px; color: #020202;">State :</label>
           <input id="State" type="text" name="State" rows="4" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;"></input>
           
           <label for="Country" style="display: block; margin-bottom: 8px; color: #020202;">Country :</label>
           <input id="Country" type="text" name="Country" rows="4" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;"></input>
   

           <label for="mobileNumber" style="display: block; margin-bottom: 8px; color: #020202;">Mobile No. <span style="color: red;"> * </span>:</label>
           <input type="tel" id="mobileNumber" name="mobileNumber" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
   
           <label for="email" style="display: block; margin-bottom: 8px; color: #020202;">Email Id <span style="color: red;"> * </span>:</label>
           <input type="email" id="email" name="email" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
   
           <label for="applicationCategory" style="display: block; margin-bottom: 8px; color: #020202;">Application Category :</label>
           <select id="applicationCategory" name="applicationCategory" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
               <option value="Painting &amp; Coating">Painting &amp; Coating</option>
               <option value="Vision Inspections">Vision Inspections</option>
               <option value="Dosing &amp; Dispensing">Dosing &amp; Dispensing</option>
               <option value="Greasing &amp; Lubrications">Greasing &amp; Lubrications</option>
               <option value="Fluid Handling">Fluid Handling</option>
               <option value="IOT, MES &amp; Data Analytics">IOT, MES &amp; Data Analytics</option>

           </select>
   
           <label for="customerType" style="display: block; margin-bottom: 8px; color: #020202;">Customer Type :</label>
           <select id="customerType" name="customerType" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
           <option value="CUSTOMER DATABASE">Customer Database</option>
           <option value="Automative OEM">Automative OEM</option>
           <option value="Tier 1 Or Tier 2">Tier 1 / Tier 2</option>
           <option value="General Industry">General Industry</option>
           
           </select>
   
           <label for="referenceFilter" style="display: block; margin-bottom: 8px; color: #020202;">Reference:</label>
           <input type="text" id="reference" name="reference" style="width: 100%; padding: 10px; margin-bottom: 12px; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px;">
   
              <center><button id="BtnClear" style=" background-color: #FF5733; color: white; cursor: pointer; padding: 12px 20px; border: none; border-radius: 4px; font-size: 16px; width : 100px;">Clear</button>
              &nbsp; &nbsp;  <button id="BtnSubmit" style=" background-color: #4caf50; color: white; cursor: pointer; padding: 12px 20px; border: none; border-radius: 4px; font-size: 16px;">Submit</button>
              </center> 
   
       </div>
       <br> <br> <br> <br> <br> <br> <br> <br> <br> <br>
   </div>
   
       </section> 
   </div>
   </div>
    </section>`;
    
    this._bindEvents();

  }


  private _bindEvents(): void {
      
 
    this.domElement.querySelector('#BtnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#BtnClear').addEventListener('click', () => { this.ClearForm(); });

  

  }

 private   getAllListItems(list) {
  try {
      let allItems = [];
      let batch =   list.items.get();
      while (batch.length > 0) {
          allItems = allItems.concat(batch);
          if (batch.hasNextChunk) {
              batch =   batch.getNextChunk();
          } else {
              break;
          }
      }
      console.log(`All items from ${list.title}:`, allItems);
      return allItems;
  } catch (error) {
      console.error(`Error getting all items from ${list.title}:`, error);
  }
}

  private  ClearForm() : void {
    var result = window.confirm("Are you sure you want to clear all fields ?");
            if (result) {
    document.getElementById("companyName")["value"] = "";
    document.getElementById("customerName")["value"] = "";
    document.getElementById("address")["value"] = "";
    document.getElementById("City")["value"] = "";
    document.getElementById("State")["value"] = "";
    document.getElementById("Country")["value"] = "";
    document.getElementById("mobileNumber")["value"] = "";
    document.getElementById("email")["value"] = "";
    document.getElementById("applicationCategory")["selectedIndex"] = 0;
    document.getElementById("customerType")["selectedIndex"] = 0;
    document.getElementById("reference")["value"] = "";
    }
  }

  private addListItem() : void {
   
   
    var companyName = document.getElementById("companyName")["value"];
    var customerName = document.getElementById("customerName")["value"];
    var address = document.getElementById("address")["value"];

    var city = document.getElementById("City")["value"];
    var state = document.getElementById("State")["value"];
    var country = document.getElementById("Country")["value"];


    var mobileNumber = document.getElementById("mobileNumber")["value"];
    var email = document.getElementById("email")["value"];
    var applicationCategory = document.getElementById("applicationCategory")["value"];
    var customerType = document.getElementById("customerType")["value"];
    var reference = document.getElementById("reference")["value"];

    
     
     

    let web = new Web ("https://cygniiautomationpvtltd.sharepoint.com/sites/SalesandMarketing2");
   
    // // const list = web.lists.getByTitle("Paint Manufacturing");
    // const list1 = web.lists.getByTitle("ADHESIVE MANUFACTURER LIST");
    // this.getAllListItems(list1);
    // const list2 = web.lists.getByTitle("Grease & Lubricants List");
    // this.getAllListItems(list2);
    // const list3 = web.lists.getByTitle("Robot Manufacturer List");
    // this.getAllListItems(list3);
 
    // const list = web.lists.getByTitle("Paint Manufacturing");
    // this.getAllListItems(list);



    if (customerType === "CUSTOMER DATABASE"){
      web.lists.getByTitle('Customer List').items.add({
        CompanyName : companyName ,  
        Name : customerName ,
        Address :address ,
        City : city ,
        STATE : state,
        Country : country,
        Mobile : mobileNumber ,
        Email : email ,
        SolutionGroup : applicationCategory ,
        Reference : reference
      }).then(r => {
                alert("CUSTOMER Record Added Successfully");
       
      });
 
    }

 
    if (customerType === "Automative OEM"){
      web.lists.getByTitle('Automative OEM').items.add({
        CompanyName : companyName ,  
        CustomerNameorPersonName : customerName ,
        Address :address ,
        City : city ,
        State : state,
        Country : country,
        MobileNo : mobileNumber ,
        EmailId : email,
        ApplicationCategory : applicationCategory ,
        Reference : reference
                 
      }).then(r => {
                alert("Automative OEM Record Added Successfully"); });
    }

    
    
    if (customerType === "Tier 1 Or Tier 2"){
      web.lists.getByTitle('Tier 1/Tier 2').items.add({
        CompanyName : companyName ,  
        CustomerNameorPersonName : customerName ,
        Address :address ,
        City : city ,
        State : state,
        Country : country,
        MobileNo : mobileNumber ,
        EmailId : email   ,
        ApplicationCategory : applicationCategory ,
        Reference : reference
         
      }).then(r => {
                alert("Tier 1 Or Tier 2 Record Added Successfully"); });
    }

    

    if (customerType === "General Industry"){
      web.lists.getByTitle('General Industry').items.add({
        CompanyName : companyName ,  
        CustomerNameorPersonName : customerName ,
        Address :address ,
        City : city ,
        State : state,
        Country : country,
        MobileNo : mobileNumber ,
        EmailId : email,
        ApplicationCategory : applicationCategory ,
        Reference : reference
                 
      }).then(r => {
                alert("General Industry Record Added Successfully"); });
    }
  
  }





  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
