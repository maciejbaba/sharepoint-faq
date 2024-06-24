import { Version, Log } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

export interface IFaqWebPartProps {
  question1: string;
  answer1: string;
  question2: string;
  answer2: string;
  question3: string;
  answer3: string;
}

export default class FaqWebPart extends BaseClientSideWebPart<IFaqWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <style>
        .faq-container {
          font-family: Arial, sans-serif;
          margin: 20px;
        }
        .faq-item {
          margin-bottom: 10px;
        }
        .faq-question {
          font-weight: bold;
          cursor: pointer;
          margin-bottom: 5px;
        }
        .faq-answer {
          display: none;
          padding-left: 20px;
        }
        .search-container {
          margin-bottom: 20px;
        }
        .search-input {
          width: 100%;
          padding: 10px;
          font-size: 16px;
        }
      </style>
      <div class="search-container">
        <input type="text" class="search-input" placeholder="Search..." />
      </div>
      <div class="faq-container">
        ${this.renderFaqItem(
          this.properties.question1,
          this.properties.answer1
        )}
        ${this.renderFaqItem(
          this.properties.question2,
          this.properties.answer2
        )}
        ${this.renderFaqItem(
          this.properties.question3,
          this.properties.answer3
        )}
      </div>
    `;

    // so basically every 500ms we want to re-render the web part
    // if this is not done, the web part will not be updated
    // because for some reason the event parts do not cause rerender
    setInterval(() => {
      this.render();
    }, 500);

    this.addEventListeners();
    console.log("FAQ web part rendered");
    this.addSearchListener();
  }

  private addSearchListener(): void {
    const searchInput = this.domElement.querySelector(
      ".search-input"
    ) as HTMLInputElement;

    if (searchInput) {
      Log.info("Search", "Search input found", this.context.serviceScope);
      searchInput.addEventListener("input", () => {
        const searchText = searchInput.value.toLowerCase();
        Log.info("Search Text", searchText, this.context.serviceScope);
        const faqItems =
          this.domElement.querySelectorAll<HTMLElement>(".faq-item");
        faqItems.forEach((faqItem) => {
          if (!searchText) {
            faqItem.style.display = "block";
            return;
          }

          const question = faqItem.querySelector(
            ".faq-question"
          ) as HTMLElement;
          const answer = faqItem.querySelector(".faq-answer") as HTMLElement;

          if (!question.textContent || !answer.textContent) {
            console.error("Question or answer not found");
            return;
          }

          if (
            question.textContent.toLowerCase().indexOf(searchText) !== -1 ||
            answer.textContent.toLowerCase().indexOf(searchText) !== -1
          ) {
            faqItem.style.display = "block";
          } else {
            faqItem.style.display = "none";
          }
        });
      });
    } else {
      console.error("Search input not found");
    }
  }

  private renderFaqItem(question: string, answer: string): string {
    return `
      <div class="faq-item">
        <div class="faq-question">${question}</div>
        <div class="faq-answer">${answer}</div>
      </div>
    `;
  }

  private addEventListeners(): void {
    const questions = this.domElement.querySelectorAll(".faq-question");
    questions.forEach((question) => {
      question.addEventListener("click", () => {
        const answer = question.nextElementSibling as HTMLElement;
        if (answer.style.display === "none" || answer.style.display === "") {
          answer.style.display = "block";
        } else {
          answer.style.display = "none";
        }
      });
    });
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure your FAQ items",
          },
          groups: [
            {
              groupName: "FAQ Item 1",
              groupFields: [
                PropertyPaneTextField("question1", {
                  label: "Question 1",
                }),
                PropertyPaneTextField("answer1", {
                  label: "Answer 1",
                  multiline: true,
                }),
              ],
            },
            {
              groupName: "FAQ Item 2",
              groupFields: [
                PropertyPaneTextField("question2", {
                  label: "Question 2",
                }),
                PropertyPaneTextField("answer2", {
                  label: "Answer 2",
                  multiline: true,
                }),
              ],
            },
            {
              groupName: "FAQ Item 3",
              groupFields: [
                PropertyPaneTextField("question3", {
                  label: "Question 3",
                }),
                PropertyPaneTextField("answer3", {
                  label: "Answer 3",
                  multiline: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
