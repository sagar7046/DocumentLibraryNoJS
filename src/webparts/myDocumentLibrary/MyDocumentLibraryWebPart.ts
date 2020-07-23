import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./MyDocumentLibraryWebPart.module.scss";
import * as strings from "MyDocumentLibraryWebPartStrings";
import { SPComponentLoader } from "@microsoft/sp-loader";
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

  GetUserList = () => {
    const array = [];
    fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/siteusers`, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json; odata=nometadata",
        "Content-Type": "application/json; odata=nometadata",
      },
    })
      .then((response) => response.json())
      .then((result) => {
        var element = document.getElementById("userList").innerHTML;
        result.value.map((item) => {
          array.push(item);
        });
        const suggestedResults = array.slice(0, 4);
        console.log(suggestedResults);
        suggestedResults.map((item) => {
          element += `<div class="ms-PeoplePicker-result" tabindex="1">
                            <div class="ms-Persona ms-Persona--xs">
                              <div class="ms-Persona-imageArea">
                                <div class="ms-Persona-initials ms-Persona-initials--blue">RM</div>
                              </div>
                              <div class="ms-Persona-presence">
                              </div>
                              <div class="ms-Persona-details">
                                <div class="ms-Persona-primaryText">${item["Title"]}</div>
                                <div class="ms-Persona-secondaryText">${item["Email"]}</div>
                              </div>
                            </div>
                            <button class="ms-PeoplePicker-resultAction">
                              <i class="ms-Icon ms-Icon--Clear"></i>
                            </button>
                        </div>`;
        });
        document.getElementById("userList").innerHTML = element;
      });
    console.log(array);
  };
  public render(): void {
    let val = this.GetUserList();
    console.log(val);
    const PeoplePicker = `<div class="ms-PeoplePicker ms-PeoplePicker--compact">
                            <div class="ms-PeoplePicker-searchBox">
                              <div class="ms-TextField ms-TextField--textFieldUnderlined">
                                <input class="ms-TextField-field" type="text" value="" placeholder="Select or enter an option">
                              </div>
                            </div>
                            <div class="ms-PeoplePicker-results ms-PeoplePicker-results--compact">
                              <div class="ms-PeoplePicker-resultGroup" id="userList">
                                <div class="ms-PeoplePicker-resultGroupTitle">
                                  Contacts
                                </div>
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
          <div class="${styles.row} ${styles.content}" >
            <div class="${styles.column}">
              <input type="file" id="inputFile" multiple hidden />
              <button class="ms-Button ms-Button--primary" id="btnFile">
                <span class="ms-Button-label">Add New File</span>
              </button>
              ${PeoplePicker}
            </div>
          </div>
        </div>
      </div>
      `;
    this.setButtonEventHandlers();
  }

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

  UploadFile(file, fileName, folder, digest) {
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
        .then((result) => console.log(result));
    };
  }

  setButtonEventHandlers = () => {
    document.getElementById("btnFile").addEventListener("click", () => {
      document.getElementById("inputFile").click();
    });
    document.querySelector("#inputFile").addEventListener("change", () => {
      try {
        const element = document.getElementById("inputFile")["files"];

        console.log(element);
        this.GetDigest()
          .then((r) => {
            for (let i = 0; i < element.length; i++) {
              var file = element[i];
              var fileName = file.name;
              var folder = "Documents";
              this.UploadFile(file, fileName, folder, r["FormDigestValue"]);
            }
          })
          .catch((err) => console.log(err));
      } catch (error) {
        console.log(error);
      }
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
