import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TestFormWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp, { sp } from 'sp-pnp-js';

import * as $ from 'jquery';
require('bootstrap');
require('./css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
require('appjs');
require('sppeoplepicker');
require('jqueryui');

export interface ITestFormWebPartProps {
  description: string;
}

export default class TestFormWebPart extends BaseClientSideWebPart<ITestFormWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div id="container" class="container">
      <div class="panel">
          <div class="panel-body">
              <div class="row">
                  <div class="col-lg-3 control-padding">
                      <label>Title</label>
                      <input type='textbox' name='txtTitle' id='txtTitle' class="form-control" value=""
                          placeholder="">
                  </div>
                  <div class="col-lg-3 control-padding">
                      <label>Person</label>

                      <div id="ppDefault"></div>
                  </div>
              </div>

              <div class="row">
                <div class="col-lg-3 control-padding">
                  <label>Time</label>
                  <div class="input-group date" data-provide="datepicker">
                      <input type="text" class="form-control" id="txtDate" name="txtDate">
                  </div>
                </div>
              </div>

              <div class="row">
                  <div class="col-lg-4 control-padding">
                      <label>Category</label>
                      <select name="ddlCategory" id="ddlCategory" class="form-control">

                      </select>
                  </div>
              </div>

              <div class="row">
                <div class="col-lg-10">
                <label>Note</label>
                <div id="txtNote">

                </div>
                </div>
              </div>

              <div class="row" style="margin-top:1rem;">
                  <div class="col col-lg-12">
                      <button type="button" class="btn btn-primary buttons" id="btnSubmit">Save</button>
                      <button type="button" class="btn btn-info buttons" id="btnUpdate">Update</button>
                      <button type="button" class="btn btn-warning buttons" id="btnDelete">Delete</button>
                  </div>
              </div>
          </div>
        </div>
        <div class="panel">
          <ul>
            <li>Form dùng để CRUD dữ liệu từ list (nếu bị lỗi F5 lại do không hiểu tại sao jquery nó lại không load trước được gây nên lỗi các thư viện khác)</li>
            <li>Update và delete theo trường Title</li>
            <li>Trường Person lấy dữ liệu account user theo tên (bắt buộc)</>
            <li>Trường Time để chọn ngày</li>
            <li>Trường Category để lấy dữ liệu từ column 'Gift name' trong list 'Gift'</li>
            <li>Trường Note để lấy dữ liệu từ list 'Power app form' column 'Multiple lines of text' và lọc theo column 'People User' đã nhập ở trường Person</li>
          </ul>
        </div>
      </div>`;

    (<any>$("#txtDate")).datepicker(
      {
        changeMonth: true,
        changeYear: true,
        dateFormat: "mm/dd/yy"
      }
    );
    (<any>$('#ppDefault')).spPeoplePicker({
      minSearchTriggerLength: 2,
      maximumEntitySuggestions: 10,
      principalType: 1,
      principalSource: 15,
      searchPrefix: '',
      searchSuffix: '',
      displayResultCount: 6,
      maxSelectedUsers: 1
    });
    this.AddEventListeners();
    this.getCategoryData();
    // this.getSubCategoryData();
  }

  private AddEventListeners(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this.SubmitData());
    document.getElementById('btnUpdate').addEventListener('click', () => this.UpdateData());
    document.getElementById('btnDelete').addEventListener('click', () => this.DeleteData());
    document.getElementById('ppDefault').addEventListener('change', () => this.getNoteData());
  }

  private SubmitData() {
    var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
    var userDetails = this.GetUserId(userinfo[0].email.toString());
    var userId = userDetails.d.Id;

    pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.add({
      Title: $("#txtTitle").val().toString(),
      PersonId: userId,
      Time: $("#txtDate").val().toString(),
      Category: $("#ddlCategory").val().toString(),
      Note: $("#txtNote").html(),
    })
    .then(() => {
      alert("Thêm ok")
    })
    .catch((ex) => {
      console.log(ex.message);
    });
  }

  private GetUserId(userName) {
    var siteUrl = this.context.pageContext.web.absoluteUrl;
    var call = $.ajax({
      url: siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",
      method: "GET",
      headers: { "Accept": "application/json; odata=verbose" },
      async: false,
      dataType: 'json'
    }).responseJSON;
    return call;
  }


  private UpdateData() {
    let title = $("#txtTitle").val().toString();
    let items = pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.top(1).filter("Title eq '"+title+"'").get()
    .then((response) => {
      if(response.length > 0) {

        var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
        var userDetails = this.GetUserId(userinfo[0].email.toString());
        var userId = userDetails.d.Id;

        const updatedItem = sp.web.lists.getByTitle("Test CRUD SPFx").items.getById(response[0].Id).update({
          PersonId: userId,
          Time: $("#txtDate").val().toString(),
          Category: $("#ddlCategory").val().toString(),
          Note: $("#txtNote").html(),
        })
        .then(() => {
          alert("update thanh cong");
        })
        .catch(() => {
          alert("something wrong wrong");
        })
      }
      else {
        alert("khong tim thay title");
      }
    });
  }

  private DeleteData() {
    let title = $("#txtTitle").val().toString();
    let items = pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.top(1).filter("Title eq '"+title+"'").get()
    .then((response) => {
      if(response.length > 0) {
        const updatedItem = sp.web.lists.getByTitle("Test CRUD SPFx").items.getById(response[0].Id).delete()
        .then(() => {
          alert("delete thanh cong");
        })
      }
      else {
        alert("khong tim thay title");
      }
    });
  }

  private _getCategoryData(): any {
    return pnp.sp.web.lists.getByTitle("Gift").items.select("GiftName").get().then((response) => {
      return response;
    });
  }

  private getCategoryData(): any {
    this._getCategoryData()
      .then((response) => {
        this._renderCategoryList(response);
      });
  }

  private _renderCategoryList(items: any): void {

    let html: string = '';
    html += `<option value="Select Category" selected>Select Category</option>`;
    items.forEach((item: any) => {
      html += `
       <option value="${item.GiftName}">${item.GiftName}</option>`;
    });
    const listContainer1: Element = this.domElement.querySelector('#ddlCategory');
    listContainer1.innerHTML = html;
  }

  private _getNoteData(): any {

    return pnp.sp.web.lists.getByTitle("Power app form").items.get().then((response) => {
      var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
      var userDetails = this.GetUserId(userinfo[0].email.toString());
      var userId = userDetails.d.Id;

      var listNote = [];

      response.forEach(items => {
        if (userId === items.People_x0020_UserId) {
          var note = {
            "note": items.Multiplelinesoftext
          };
          listNote.push(note);
        }
      });
      return listNote;
    });
  }

  private getNoteData(): any {
    this._getNoteData()
      .then((response) => {
        this._renderNoteList(response);
      });
  }

  private _renderNoteList(items: any): void {

    let html: string = '';
    items.forEach((item: any) => {
      html += item.note;
    });
    const listContainer1: Element = this.domElement.querySelector('#txtNote');
    listContainer1.innerHTML = html;
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

