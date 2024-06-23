import * as React from "react";
import styles from "./Faq.module.scss";
import type { IFaqProps } from "./IFaqProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import "./reactAccordion.css";

import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from "react-accessible-accordion";

export interface IReactAccordionState {
  items: Array<any>;
  allowMultipleExpanded: boolean;
  allowZeroExpanded: boolean;
}

export default class Faq extends React.Component<
  IFaqProps,
  IReactAccordionState
> {
  constructor(props: IFaqProps) {
    super(props);

    this.state = {
      items: [],
      allowMultipleExpanded: this.props.allowMultipleExpanded,
      allowZeroExpanded: this.props.allowZeroExpanded,
    };
    this.getListItems();
  }

  private getListItems(): void {
    const { listId } = this.props;
    if (listId) {
      sp.web.lists
        .getById(listId)
        .items.select("Title", "Content")
        .get()
        .then((results: Array<any>) => {
          this.setState({ items: results });
        })
        .catch((error) => {
          console.error("Failed to get list items!", error);
        });
    }
  }

  public componentDidUpdate(prevProps: IFaqProps): void {
    if (prevProps.listId !== this.props.listId) {
      this.getListItems();
    }

    if (
      prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded ||
      prevProps.allowZeroExpanded !== this.props.allowZeroExpanded
    ) {
      this.setState({
        allowMultipleExpanded: this.props.allowMultipleExpanded,
        allowZeroExpanded: this.props.allowZeroExpanded,
      });
    }
  }

  public render(): React.ReactElement<IFaqProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      listId,
      onConfigure,
      displayMode,
      accordionTitle,
      updateProperty,
    } = this.props;
    const { allowMultipleExpanded, allowZeroExpanded, items } = this.state;
    const listSelected = !!listId;

    return (
      <section
        className={`${styles.faq} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is an extensibility model for
            Microsoft Viva, Microsoft Teams, and SharePoint. It&#39;s the
            easiest way to extend Microsoft 365 with automatic Single Sign On,
            automatic hosting, and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li>
              <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
                SharePoint Framework Overview
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-graph"
                target="_blank"
                rel="noreferrer"
              >
                Use Microsoft Graph in your solution
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-teams"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Teams using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-viva"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Viva Connections using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-store"
                target="_blank"
                rel="noreferrer"
              >
                Publish SharePoint Framework applications to the marketplace
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-api"
                target="_blank"
                rel="noreferrer"
              >
                SharePoint Framework API reference
              </a>
            </li>
            <li>
              <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
                Microsoft 365 Developer Community
              </a>
            </li>
          </ul>
        </div>
        <div className={styles.reactAccordion}>
          {!listSelected && (
            <Placeholder
              iconName="MusicInCollectionFill"
              iconText="Configure your web part"
              description="Select a list with a Title field and Content field to have its items rendered in a collapsible accordion format"
              buttonLabel="Choose a List"
              onConfigure={onConfigure}
            />
          )}
          {listSelected && (
            <div>
              <WebPartTitle
                displayMode={displayMode}
                title={accordionTitle}
                updateProperty={updateProperty}
              />
              <Accordion
                allowZeroExpanded={allowZeroExpanded}
                allowMultipleExpanded={allowMultipleExpanded}
              >
                {items.map((item) => (
                  <AccordionItem key={item.Id}>
                    <AccordionItemHeading>
                      <AccordionItemButton>{item.Title}</AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <p dangerouslySetInnerHTML={{ __html: item.Content }} />
                    </AccordionItemPanel>
                  </AccordionItem>
                ))}
              </Accordion>
            </div>
          )}
        </div>
      </section>
    );
  }
}
