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
    this.domElement.innerHTML =
    `
      <div id="container">
        <h2>Test CRUD SPFx list</h2>
        <button type="button" class="btn btn-primary buttons btn-lg" data-toggle="modal" data-target="#myModal">New</button>
        <!-- *********************************************************** -->
        <div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
          <div class="modal-dialog" role="document">
            <div class="modal-content">
              <div class="modal-header">
                <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <h4 class="modal-title" id="myModalLabel">New</h4>
              </div>
              <div class="modal-body">
                <div class="row">
                  <div class="col-lg-5 control-padding">
                      <label>Title</label>
                      <input type='textbox' name='txtTitle' id='txtTitle' class="form-control" value=""
                          placeholder="">
                  </div>
                  <div class="col-lg-5 control-padding">
                      <label>Person</label>

                      <div id="ppDefault"></div>
                  </div>
                </div>

                <div class="row">
                  <div class="col-lg-5 control-padding">
                    <label>Time</label>
                    <div class="input-group date" data-provide="datepicker">
                        <input type="text" class="form-control" id="txtDate" name="txtDate">
                    </div>
                  </div>
                </div>

                <div class="row">
                    <div class="col-lg-5 control-padding">
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
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                <button type="button" class="btn btn-primary" id="btnSubmit">Save</button>
              </div>
            </div>
          </div>
        </div>
        <!-- *********************************************************** -->
        <div class="modal fade" id="modalUpdate" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
          <div class="modal-dialog" role="document">
            <div class="modal-content">
              <div class="modal-header">
                <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <h4 class="modal-title" id="myModalLabel">Update</h4>
              </div>
              <div class="modal-body">
                <div class="row">
                  <div class="col-lg-5 control-padding">
                      <label>Title</label>
                      <input type='textbox' name='txtTitle1' id='txtTitle1' class="form-control" value=""
                          placeholder="">
                  </div>
                  <div class="col-lg-5 control-padding">
                      <label>Person</label>

                      <div id="ppDefault1"></div>
                  </div>
                </div>

                <div class="row">
                  <div class="col-lg-5 control-padding">
                    <label>Time</label>
                    <div class="input-group date" data-provide="datepicker">
                        <input type="text" class="form-control" id="txtDate1" name="txtDate1">
                    </div>
                  </div>
                </div>

                <div class="row">
                    <div class="col-lg-5 control-padding">
                        <label>Category</label>
                        <select name="ddlCategory1" id="ddlCategory1" class="form-control">

                        </select>
                    </div>
                </div>

                <div class="row">
                  <div class="col-lg-10">
                  <label>Note</label>
                  <div id="txtNote1">

                  </div>
                  </div>
                </div>
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                <button type="button" class="btn btn-primary" id="btnUpdate">Save</button>
              </div>
            </div>
          </div>
        </div>
        <!-- *********************************************************** -->
          <div class="panel">
            <table class="table table-striped table-hover table-bordered">
              <thead>
                <tr>
                  <th class="info">Title</th>
                  <th class="success">Person</th>
                  <th class="warning">Time</th>
                  <th class="danger">Category</th>
                  <th class="active">Note</th>
                  <th class="info">Action</th>
                </tr>
              </thead>
              <tbody id="tableData">

              </tbody>
            </table>
          </div>
      </div>
    `;

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
    (<any>$('#ppDefault1')).spPeoplePicker({
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
    this.ReadData();

  }

  private AddEventListeners(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this.SubmitData());
    document.getElementById('btnUpdate').addEventListener('click', () => this.UpdateData());
    // document.getElementById('btnModalUpdate').addEventListener('click', () => this.showModalUpdateData(this));
    // document.getElementById('btnModalDelete').addEventListener('click', () => this.showModalDeleteData(this));
    document.getElementById('ppDefault').addEventListener('change', () => this.getNoteData());
    document.getElementById('ppDefault1').addEventListener('change', () => this.getNoteData1());
  }

  private ReadData() {
    pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.get()
    .then((response) => {
      response
      let html: string = '';
      response.forEach((item: any) => {
        html += `
          <tr>
            <td>${item.Title}</td>
            <td>${item.PersonId}</td>
            <td>${this.dateFormat(item.Time)}</td>
            <td>${item.Category}</td>
            <td>${item.Note}</td>
            <td>
              <button class="btn btn-info buttons btnModalUpdate" data-title='${item.Title}' data-personId='${item.PersonId}' data-date='${this.dateFormat(item.Time)}' data-category='${item.Category}' type="button">Update</button>
              <button class="btn btn-warning buttons btnModalDelete" data-title='${item.Title}' type="button">Delete</button>
            </td>
          </tr>
        `;
      });
      const table: Element = this.domElement.querySelector('#tableData');
      table.innerHTML = html;
      this.setEventUpdateDeleteButton();
    })
    .catch((ex) => {
      console.log(ex.message);
    });
  }

  private setEventUpdateDeleteButton() {
    let btnUpdate = document.getElementsByClassName('btnModalUpdate');
    for(let i = 0; i < btnUpdate.length; i++) {
      btnUpdate[i].addEventListener("click", () => {
        this.showModalUpdateData(btnUpdate[i])
      }, false)
    }

    let btnDelete = document.getElementsByClassName('btnModalDelete');
    for(let i = 0; i < btnDelete.length; i++) {
      btnDelete[i].addEventListener("click", () => {
        this.DeleteData(btnDelete[i].getAttribute('data-title'));
      }, false)
    }
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
      alert("Thêm ok");
      location.reload();
    })
    .catch((ex) => {
      alert("Something wrong wong");
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
    let title = $("#txtTitle1").val().toString();
    let items = pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.top(1).filter("Title eq '"+title+"'").get()
    .then((response) => {
      if(response.length > 0) {

        var userinfo = (<any>$('#ppDefault1')).spPeoplePicker('get');
        var userDetails = this.GetUserId(userinfo[0].email.toString());
        var userId = userDetails.d.Id;

        const updatedItem = sp.web.lists.getByTitle("Test CRUD SPFx").items.getById(response[0].Id).update({
          PersonId: userId,
          Time: $("#txtDate1").val().toString(),
          Category: $("#ddlCategory1").val().toString(),
          Note: $("#txtNote1").html(),
        })
        .then(() => {
          alert("update thanh cong");
          location.reload();
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

  private showModalUpdateData(data : any) {
    let title = data.getAttribute('data-title');
    let date = data.getAttribute('data-date');
    let category = data.getAttribute('data-category');

    (<any>$('#modalUpdate')).modal('show');
    $('#modalUpdate').on('show.bs.modal', function () {
       $('#txtTitle1').val(title);
       $('#txtDate1').val(date);
       $('#ddlCategory1').val(category);

    })
  }

  private DeleteData(title : any) {
    let items = pnp.sp.web.lists.getByTitle("Test CRUD SPFx").items.top(1).filter("Title eq '"+title+"'").get()
    .then((response) => {
      if(response.length > 0) {
        const updatedItem = sp.web.lists.getByTitle("Test CRUD SPFx").items.getById(response[0].Id).delete()
        .then(() => {
          alert("delete thanh cong");
          location.reload();
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
    const listContainer2: Element = this.domElement.querySelector('#ddlCategory1');
    listContainer2.innerHTML = html;
  }

  //note create
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

  //note update
  private _getNoteData1(): any {

    return pnp.sp.web.lists.getByTitle("Power app form").items.get().then((response) => {
      var userinfo = (<any>$('#ppDefault1')).spPeoplePicker('get');
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

  private getNoteData1(): any {
    this._getNoteData1()
      .then((response) => {
        this._renderNoteList1(response);
      });
  }

  private _renderNoteList1(items: any): void {

    let html: string = '';
    items.forEach((item: any) => {
      html += item.note;
    });
    const listContainer1: Element = this.domElement.querySelector('#txtNote1');
    listContainer1.innerHTML = html;
  }

  private dateFormat(dateString) {
    let date = new Date(dateString);
    return date.getFullYear() + '-' + (date.getMonth() + 1) + '-' + date.getDate();
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

