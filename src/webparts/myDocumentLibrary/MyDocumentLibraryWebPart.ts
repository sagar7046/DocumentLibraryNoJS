import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape, sortBy } from "@microsoft/sp-lodash-subset";

import styles from "./MyDocumentLibraryWebPart.module.scss";
import * as strings from "MyDocumentLibraryWebPartStrings";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPUpdate } from "./SPOHelper";
export interface IMyDocumentLibraryWebPartProps {
  description: string;
}

export default class MyDocumentLibraryWebPart extends BaseClientSideWebPart<
  IMyDocumentLibraryWebPartProps
> {
  //MyDocumentLibraryWebPart.data.push(result.value);
  data = new Array();
  public SelectFile() {
    document.getElementById("inputFile").click();
  }
  constructor() {
    super();
    SPComponentLoader.loadCss(
      "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css"
    );
    SPComponentLoader.loadCss(
      "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css"
    );
    SPComponentLoader.loadScript(
      "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"
    ).then(() => {
      var PeoplePickerElements = document.querySelectorAll(".ms-PeoplePicker");
      for (var i = 0; i < PeoplePickerElements.length; i++) {
        //@ts-ignore
        new fabric["PeoplePicker"](PeoplePickerElements[i]);
      }
    });
  }

  public render(): void {
    const PeoplePicker = `<div class="ms-PeoplePicker ms-PeoplePicker--compact">
                            <div class="ms-PeoplePicker-searchBox">
                              <div class="ms-TextField ms-TextField--textFieldUnderlined">
                                <input class="ms-TextField-field" type="text" id="PeoplePicker" value="" placeholder="Select or enter an option" autocomplete="off">
                              </div>
                            </div>
                            <div class="ms-PeoplePicker-results ms-PeoplePicker-results--compact">
                              <div class="ms-PeoplePicker-resultGroup" id="userList">
                            </div>
                          </div>`;

    this.domElement.innerHTML = `
      <div class="${styles.myDocumentLibrary}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <h1>${strings.Heading}</h1>
            </div>
          </div>
          <div class="${styles.row} ${styles.content}">
            <div class="${styles.column}">
              <input type="file" id="inputFile" multiple hidden />
              <button class="ms-Button ms-Button--primary" id="btnFile">
                <span class="ms-Button-label">Add New File</span>
              </button>
              <span id="files"></span><br/>
            </div>
            <div class="${styles.column}">
             ${PeoplePicker}
            </div>
            <div class="${styles.column}">
              <button class="ms-Button ms-Button--primary" id="btnPeople">
                <span class="ms-Button-label">Upload Files</span>
              </button>
              <button class="ms-Button ms-Button--primary" id="btnPeople">
                <span class="ms-Button-label" id="view">View Files</span>
              </button>
            </div>
            <div class="${styles.column}" id="viewTable">
              <div class="${styles.loadingIcon}" id="loading">
                <img src="https://gayasagar.sharepoint.com/sites/MySite/Images1/loading.gif" height="40" width="40"></img>
              </div>
            </div>
          </div>
        </div>
      </div>
      `;
    this.setButtonEventHandlers();
  }

  //method to create table
  DisplayTable = () => {
    const element = `
    <table class="ms-Table">
    <thead>
      <tr>
        <th>FileName</th>
        <th>Modified</th>
        <th>Action</th>
      </tr>
    </thead>
    <tbody id="data">
    </tbody>
</table>`;
    return element;
  };

  //Fill table data
  FillTable = () => {
    fetch(
      `${this.context.pageContext.web.absoluteUrl}/_api/Web/Lists/getByTitle('Documents')/RootFolder/Files/`,
      {
        method: "GET",
        credentials: "include",
        headers: {
          Accept: "application/json; odata=nometadata",
          "Content-Type": "application/json; odata=nometadata",
        },
      }
    )
      .then((res) => res.json())
      .then((r) => {
        var el = "";
        console.log(r);
        document.getElementById("viewTable").innerHTML = this.DisplayTable();
        console.log(r.value);
        r.value.map((item) => {
          el += `<tr>
        <td>${item.Name}</td>
        <td>${item.TimeLastModified}</td>
        <td>
        <button class="deletefiles ${styles.tableData}" value="${item.Name}"><img src="https://gayasagar.sharepoint.com/sites/MySite/Images1/icons8-delete-64.png" height="25" width="25"></img></button>
        <button class="downloadFile ${styles.tableData}" value="${item.ServerRelativeUrl}"><img src="https://gayasagar.sharepoint.com/sites/MySite/Images1/icons8-download-96.png" height="25" width="25"></img></button>
        </td>
        </tr>`;
        });
        document.getElementById("data").innerHTML = el;
        document.querySelectorAll(".deletefiles").forEach((item) =>
          item.addEventListener("click", () => {
            let confirmation = confirm(
              "Are you sure you want to delete this item?"
            );
            confirmation === true ? this.DeleteFile(item["value"]) : null;
          })
        );
        document.querySelectorAll(".downloadFile").forEach((item) =>
          item.addEventListener("click", () => {
            let name = item["value"].split("/");
            let source = `https://gayasagar.sharepoint.com${item["value"]}`;
            let uri = encodeURI(source);
            var link = document.createElement("a");
            link.href = uri;
            link.setAttribute("download", `${name[4]}`);
            document.body.appendChild(link);
            link.click();
          })
        );
      });
  };

  //method to delete records...
  DeleteFile = (name) => {
    let fileUrl = `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents/${name}`;
    let endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrl}')`;
    this.GetDigest().then((r) => {
      fetch(`${endpoint}`, {
        method: "POST",
        headers: {
          credentials: "include",
          Accept: "application/json; odata=verbose",
          "Content-Type": "application/json; odata=verbose",
          "X-RequestDigest": r["FormDigestValue"],
          "X-HTTP-METHOD": "DELETE",
        },
      }).then((r) => {
        if (r.status === 200) {
          alert("Successfully Deleted");
          this.FillTable();
        } else {
          alert("Problem in delete look for console log for help");
          console.log(r);
        }
      });
    });
  };

  //method to get the request Digest which will be used while adding or updating record
  GetDigest() {
    return new Promise((resolve, reject) => {
      fetch(`${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
        method: "POST",
        body: null,
        headers: {
          credentials: "include", // credentials: 'same-origin',
          Accept: "application/json; odata=nometadata",
          "Content-Type": "application/json; odata=nometadata",
        },
      })
        .then((res) => res.json())
        .then((r) => resolve(r))
        .catch((err) => reject(err));
    });
  }
  //method to render the list item for people picker for each user list.
  SuggestedResults = (item: any[]) => {
    const avt = item["Title"].split(" ");
    let op = "";
    avt.map((i) => {
      op += i.substring(0, 1);
    });
    const element = `<div class="ms-PeoplePicker-result" tabindex="1">
                        <div class="ms-Persona ms-Persona--xs">
                          <div class="ms-Persona-imageArea">
                            <div class="ms-Persona-initials ms-Persona-initials--blue">${op}</div>
                          </div>
                          <div class="ms-Persona-presence">
                          </div>
                          <div class="ms-Persona-details">
                            <div class="ms-Persona-primaryText">${item["Title"]}</div>
                              <div class="ms-Persona-secondaryText">${item["Title"]}</div>
                                <input type="hidden" value ="${item["Id"]}" class="hidden-values"/>
                              </div>
                            </div>
                          <button class="ms-PeoplePicker-resultAction">
                              <i class="ms-Icon ms-Icon--Clear"></i>
                            </button>
                        </div>
                      `;
    return element;
  };

  //Method to get the users list and binding it to People Picker control
  GetUserList = (val = null) => {
    const array = [];
    const endpointUrl =
      val !== null
        ? `${this.context.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=startswith(Title,'${val}')`
        : `${this.context.pageContext.web.absoluteUrl}/_api/web/siteusers`;
    fetch(endpointUrl, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json; odata=nometadata",
        "Content-Type": "application/json; odata=nometadata",
      },
    })
      .then((response) => response.json())
      .then((result) => {
        document.getElementById("userList");
        var element =
          val === null || val === ""
            ? ` <div class="ms-PeoplePicker-resultGroupTitle">Frequent Contacts</div>`
            : `<div class="ms-PeoplePicker-resultGroupTitle">Filtered Contacts</div>`;
        result.value.map((item) => {
          array.push(item);
        });
        const suggestedResults = array.filter(
          (item) => item.PrincipalType === 1 || item.PrincipalType === 4
        );
        suggestedResults.slice(0, 3).map((item) => {
          const avt = item["Title"].split(" ");
          let op = "";
          avt.map((i) => {
            op += i.substring(0, 1);
          });
          element += this.SuggestedResults(item);
        });
        document.getElementById("userList").innerHTML = element;
      });
  };

  //method to upload file.
  UploadFile(file, fileName, folder, digest, user) {
    return new Promise((resolve, reject) => {
      var reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = (event) => {
        var siteurl = `${this.context.pageContext.web.absoluteUrl}/_api/Web/Lists/getByTitle('${folder}')/RootFolder/Files/Add(url='${fileName}',overwrite=true)`;
        fetch(`${siteurl}`, {
          method: "POST",
          headers: {
            credentials: "include",
            Accept: "application/json; odata=verbose",
            "Content-Type": "application/json; odata=verbose",
            "X-RequestDigest": digest,
          },
          body: new Blob([event.target["result"]]),
        })
          .then((response) => response.json())
          .then((result) => {
            const uri = result.d.ListItemAllFields.__deferred.uri;
            SPUpdate({
              url: `${uri}`,
              payload: { ApprovelById: user },
            })
              .then((r) => {
                resolve(r);
              })
              .catch((err) => reject(err));
          });
      };
    });
  }

  //Binding events to html element
  setButtonEventHandlers = () => {
    this.GetUserList();
    document.getElementById("btnFile").addEventListener("click", () => {
      document.getElementById("inputFile").click();
    });

    document
      .getElementById("PeoplePicker")
      .addEventListener("input", (event) => {
        console.log(event.target["value"]);
        this.GetUserList(event.target["value"]);
      });
    document.getElementById("inputFile").addEventListener("change", () => {
      const element = document.getElementById("inputFile")["files"];
      let label = "";
      for (let i = 0; i < element.length; i++) {
        var file = element[i];
        var fileName = file.name;
        document.getElementById("files").innerText += fileName + "   ";
      }
    });

    document.getElementById("btnPeople").addEventListener("click", () => {
      var id;
      const elementq = document.getElementsByClassName(
        "ms-PeoplePicker-searchBox"
      )[0];
      elementq
        .querySelectorAll(".hidden-values")
        .forEach((item) => (id = item["value"]));

      //Calling file upload method and passing parameters to it when user id is not equal to null.
      const element = document.getElementById("inputFile")["files"];
      if (id != null && element.length !== 0) {
        const itemUploaded = [];
        try {
          this.GetDigest()
            .then((r) => {
              for (let i = 0; i < element.length; i++) {
                var file = element[i];
                var fileName = file.name;

                var folder = "Documents";
                this.UploadFile(
                  file,
                  fileName,
                  folder,
                  r["FormDigestValue"],
                  id
                ).then((res) => {
                  itemUploaded.push(res);
                  if (itemUploaded.length === element.length) {
                    alert("File has been successfully uploaded.");
                  }
                });
              }
            })
            .catch((err) => console.log(err));
        } catch (error) {
          console.log(error);
        }
      } else {
        alert("Please select user or group from PeoplePicker");
      }
    });

    document.getElementById("view").addEventListener("click", () => {
      document.getElementById("loading").style.display = "flex";
      // document.getElementById("viewTable").innerHTML = this.DisplayTable();
      this.FillTable();
    });
  };

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
