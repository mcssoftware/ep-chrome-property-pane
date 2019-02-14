import * as React from "react";
import { Async } from "office-ui-fabric-react/lib/Utilities";
import { PrimaryButton, DefaultButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { Spinner, SpinnerType } from "office-ui-fabric-react/lib/Spinner";
import { Label } from "office-ui-fabric-react/lib/Label";
import TermPicker from "./TermPicker";
import { IPropertyFieldNewsSelectorHostProps, IPropertyFieldNewsSelectorHostState } from "./IPropertyFieldNewsSelectorHost";
import { ITermStore, ITerm, ISPTermStorePickerService } from "../../../services/ISPTermStorePickerService";
import styles from "./PropertyFieldNewsSelectorHost.module.scss";
import { sortBy, uniqBy, cloneDeep } from "@microsoft/sp-lodash-subset";
import TermGroup from "./TermGroup";
import FieldErrorMessage from "../../errorMessage/FieldErrorMessage";
import { IPickerTerms, IPickerTerm } from "../termStoreEntity";
import { ActiveDisplayModeType, IPropertyFieldNewsSelectorData, getPropertyFieldDefaultValue } from "../IPropertyFieldNewsSelector";
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react/lib-es2015/ChoiceGroup";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib-es2015/Dropdown";
import { ISPService } from "../../../services/ISPService";
import Header from "../../header/header";
import { initGlobalVars } from "../../../common/ep";

/**
 * Image URLs / Base64
 */
export const COLLAPSED_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA8AAAAUCAYAAABSx2cSAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAABh0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjEwcrIlkgAAAIJJREFUOE/NkjEKwCAMRdu7ewZXJ/EqHkJwE9TBCwR+a6FLUQsRwYBTeD8/35wADnZVmPvY4OOYO3UNbK1FKeUWH+fRtK21hjEG3vuhQBdOKUEpBedcV6ALExFijJBSIufcFBjCVSCEACEEqpNvBmsmT+3MTnvqn/+O4+1vdtv7274APmNjtuXVz6sAAAAASUVORK5CYII="; // /_layouts/15/images/MDNCollapsed.png
export const EXPANDED_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA8AAAAUCAYAAABSx2cSAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAABh0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjEwcrIlkgAAAFtJREFUOE9j/P//PwPZAKSZXEy2RrCLybV1CGjetWvX/46ODqBLUQOXoJ9BGtXU1MCYJM0wjZGRkaRpRtZIkmZ0jSRpBgUOzJ8wmqwAw5eICIb2qGYSkyfNAgwAasU+UQcFvD8AAAAASUVORK5CYII="; // /_layouts/15/images/MDNExpanded.png
export const GROUP_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAC9SURBVDhPY2CgNXh1qEkdiJ8D8X90TNBuJM0V6IpBhoHFgIxebKYTIwYzAMNpxGhGdsFwNoBgNEFjAWsYgOSKiorMgPgbEP/Hgj8AxXpB0Yg1gQAldYuLix8/efLkzn8s4O7du9eAan7iM+DV/v37z546der/jx8/sJkBdhVOA5qbm08ePnwYrOjQoUOkGwDU+AFowLmjR4/idwGukAYaYAkMgxfPnj27h816kDg4DPABoAI/IP6DIxZA4l0AOd9H3QXl5+cAAAAASUVORK5CYII="; // /_layouts/15/Images/EMMGroup.png
export const TERMSET_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACaSURBVDhPrZLRCcAgDERdpZMIjuQA7uWH4CqdxMY0EQtNjKWB0A/77sxF55SKMTalk8a61lqCFqsLiwKac84ZRUUBi7MoYHVmAfjfjzE6vJqZQfie0AcwBQVW8ATi7AR7zGGGNSE6Q2cyLSPIjRswjO7qKhcPDN2hK46w05wZMcEUIG+HrzzcrRsQBIJ5hS8C9fGAPmRwu/9RFxW6L8CM4Ry8AAAAAElFTkSuQmCC"; // /_layouts/15/Images/EMMTermSet.png

/**
 * Renders the controls for PropertyFieldTermPicker component
 */
export default class PropertyFieldNewsSelectorHost extends React.Component<IPropertyFieldNewsSelectorHostProps, IPropertyFieldNewsSelectorHostState> {
  private async: Async;
  private delayedValidate: (value: IPropertyFieldNewsSelectorData) => void;
  private termsService: ISPTermStorePickerService;
  private spService: ISPService;
  private previousValues: IPropertyFieldNewsSelectorData = getPropertyFieldDefaultValue();
  private cancel: boolean = true;

  /**
   * Constructor method
   */
  constructor(props: IPropertyFieldNewsSelectorHostProps) {
    super(props);
    initGlobalVars();
    this.termsService = props.termService;
    this.spService = props.spService;
    const activeValues = typeof this.props.initialValues !== "undefined" ? this.props.initialValues : getPropertyFieldDefaultValue();
    this.state = {
      activeValues,
      termStores: [],
      termStoreLoaded: false,
      openPanel: false,
      errorMessage: this.validateInternal(activeValues),
      pageDropDownOptions: this.getEmptyDropDownOption(),
      pagesLoaded: true,
    };    
    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.onSave = this.onSave.bind(this);
    this.termsChanged = this.termsChanged.bind(this);
    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.termsFromPickerChanged = this.termsFromPickerChanged.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
    this.onDisplayModeChange = this.onDisplayModeChange.bind(this);
    this.onArticleDropDownChanged = this.onArticleDropDownChanged.bind(this);
    this.onRenderDropDownTitle = this.onRenderDropDownTitle.bind(this);
  }

  /**
   * Loads the list from SharePoint current web site
   */
  private loadTermStores(): void {
    this.termsService.getTermStores().then((response: ITermStore[]) => {
      // Check if a response was retrieved
      if (response !== null) {
        this.setState({
          termStores: response,
          termStoreLoaded: true
        });
      } else {
        this.setState({
          termStores: [],
          termStoreLoaded: true
        });
      }
    });
  }

  /**
  * Loads list items from Sharepoint pages library
  *
  * @private
  * @memberof PropertyFieldNewsSelectorHost
  */
  private loadPages(value: IPropertyFieldNewsSelectorData): void {
    if (this.isvalidateNewsChannel(value)) {
      this.setState({ pagesLoaded: false || this.state.openPanel, pageDropDownOptions: this.getEmptyDropDownOption() });
      const now: string = new Date(Date.now()).toISOString();
      const camlQuery: string = "<View><Query><Where>" +
        `<And><And><Lt><FieldRef Name='ArticleStartDate' /><Value Type='DateTime'>${now}</Value></Lt>` +
        `<Gt><FieldRef Name='News_x0020_Article_x0020_Expiration_x0020_Date' /><Value Type='DateTime'>${now}</Value></Gt></And>` +
        `<Eq><FieldRef Name='News_x0020_Channel' /><Value Type='TaxonomyFieldType'>${value.NewsChannel[0].name}</Value></Eq>` +
        "</And></Where></Query><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /></ViewFields><RowLimit>200</RowLimit></View>";
      this.spService.getListItemsCaml(camlQuery, "Pages", window.Epmodern.urls.newsUrl).then((pages) => {
        const pageDropDownOptions: IDropdownOption[] = this.getEmptyDropDownOption();
        pages.filter((f) => typeof f.Title === "string" && f.Title.trim().length > 0).forEach(f => {
          pageDropDownOptions.push({
            key: f.Id.toString(),
            text: f.Title.trim()
          });
        });
        this.setState({ pagesLoaded: true, pageDropDownOptions });
      });
    } else {
      this.setState({ pageDropDownOptions: this.getEmptyDropDownOption() });
    }
  }

  private getEmptyDropDownOption(): IDropdownOption[] {
    return [{ key: "0", text: "Select Page" }];
  }

  /**
   * Validates the new custom field value
   */
  private validate(value: IPropertyFieldNewsSelectorData): void {
    const internalResult: string = this.validateInternal(value);
    if (internalResult.length < 1) {
      if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
        this.notifyAfterValidate(this.props.initialValues, value);
        return;
      }
    }

    const result: string | PromiseLike<string> = internalResult.length > 0 ? internalResult : this.props.onGetErrorMessage(value || getPropertyFieldDefaultValue());
    if (typeof result !== "undefined") {
      if (typeof result === "string") {
        if (result === "") {
          this.notifyAfterValidate(this.props.initialValues, value);
        }
        this.setState({
          errorMessage: result
        });
      } else {
        result.then((errorMessage: string) => {
          if (typeof errorMessage === "undefined" || errorMessage === "") {
            this.notifyAfterValidate(this.props.initialValues, value);
          }
          this.setState({
            errorMessage: errorMessage
          });
        });
      }
    } else {
      this.notifyAfterValidate(this.props.initialValues, value);
    }
  }

  private validateInternal(value: IPropertyFieldNewsSelectorData): string {
    if (typeof value === "undefined" || value === null) {
      return "";
    }
    if (!this.isvalidateNewsChannel(value)) {
      return "Please select news channel";
    }
    const articleIdValid: boolean = typeof value.ArticleId !== "undefined" && value.ArticleId !== null && !isNaN(parseFloat(value.ArticleId.toString())) && isFinite(value.ArticleId);
    if ((value.ActiveDisplayMode === ActiveDisplayModeType.Specific) && !articleIdValid) {
      return "Please select article";
    }
    return "";
  }

  private isvalidateNewsChannel(value: IPropertyFieldNewsSelectorData): boolean {
    const newsChannel = value.NewsChannel;
    if (typeof newsChannel === "undefined" || newsChannel === null || !Array.isArray(newsChannel) || newsChannel.length < 1) {
      return false;
    }
    return true;
  }

  /**
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: IPropertyFieldNewsSelectorData, newValue: IPropertyFieldNewsSelectorData) {
    if (this.props.onPropertyChange && newValue !== null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
      // Trigger the apply button
      if (typeof this.props.onChange !== "undefined" && this.props.onChange !== null) {
        this.props.onChange(this.props.targetProperty, newValue);
      }
    }
  }


  /**
   * Open the right Panel
   */
  private onOpenPanel(): void {
    if (this.props.disabled === true) {
      return;
    }

    // Store the current code value
    this.previousValues = cloneDeep(this.state.activeValues);
    this.cancel = true;

    this.loadTermStores();

    this.setState({
      openPanel: true,
      termStoreLoaded: false
    });
  }

  /**
   * Close the panel
   */
  private onClosePanel(): void {

    this.setState(() => {
      const newState: IPropertyFieldNewsSelectorHostState = {
        openPanel: false,
        termStoreLoaded: false
      } as IPropertyFieldNewsSelectorHostState;

      // Check if the property has to be reset
      if (this.cancel) {
        newState.activeValues = this.previousValues;
      }
      return newState;
    });
  }

  /**
   * On save click action
   */
  private onSave(): void {
    this.cancel = false;
    this.delayedValidate(this.state.activeValues);
    this.onClosePanel();
  }

  /**
   * Clicks on a node
   * @param node
   */
  private termsChanged(term: ITerm, termGroup: string, checked: boolean): void {

    let activeValues = cloneDeep(this.state.activeValues);
    if (typeof term === "undefined" || term === null) {
      return;
    }

    // Term item to add to the active nodes array
    const termItem: IPickerTerm = {
      name: term.Name,
      key: term.Id,
      path: term.PathOfTerm,
      termSet: term.TermSet.Id,
      termGroup: termGroup,
      labels: term.Labels
    };

    // Check if the term is checked or unchecked
    if (checked) {
      // Check if it is allowed to select multiple terms
      if (this.props.allowMultipleSelections) {
        // Add the checked term
        activeValues.NewsChannel.push(termItem);
        // Filter out the duplicate terms
        activeValues.NewsChannel = uniqBy(activeValues.NewsChannel, "key");
      } else {
        // Only store the current selected item
        activeValues.NewsChannel = [termItem];
      }
    } else {
      // Remove the term from the list of active nodes
      activeValues.NewsChannel = activeValues.NewsChannel.filter(item => item.key !== term.Id);
    }
    // Sort all active nodes
    activeValues.NewsChannel = sortBy(activeValues.NewsChannel, "path");
    activeValues.ArticleId = 0;
    activeValues.ActiveDisplayMode = ActiveDisplayModeType.Latest;
    // Update the current state
    this.setState({
      activeValues: activeValues
    });
    this.loadPages(activeValues);
  }

  /**
  * Fires When Items Changed in TermPicker
  * @param node
  */
  private termsFromPickerChanged(terms: IPickerTerms): void {
    const activeValues = cloneDeep(this.state.activeValues);
    activeValues.NewsChannel = terms;
    activeValues.ArticleId = 0;
    activeValues.ActiveDisplayMode = ActiveDisplayModeType.Latest;
    this.delayedValidate(activeValues);

    this.setState({
      activeValues
    });
    this.loadPages(activeValues);
  }

  private onDisplayModeChange(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void {
    if (typeof option !== "undefined") {
      let mode: ActiveDisplayModeType = ActiveDisplayModeType.Latest;
      if (option.key === ActiveDisplayModeType.Specific.toString()) {
        mode = ActiveDisplayModeType.Specific;
      }
      const data = cloneDeep(this.state.activeValues);
      data.ActiveDisplayMode = mode;
      this.delayedValidate(data);

      this.setState({ activeValues: data });
    }
  }

  private onArticleDropDownChanged(option: IDropdownOption, index?: number): void {
    const key: number = option.key as number;
    const data = cloneDeep(this.state.activeValues);
    data.ArticleId = key;
    this.delayedValidate(data);
    this.setState({ activeValues: data });
  }

  // /**
  //  * Gets the given node position in the active nodes collection
  //  * @param node
  //  */
  // private getSelectedNodePosition(node: IPickerTerm): number {
  //   for (let i = 0; i < this.state.activeNodes.length; i++) {
  //     if (node.key === this.state.activeNodes[i].key) {
  //       return i;
  //     }
  //   }
  //   return -1;
  // }

  /**
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    if (typeof this.async !== "undefined") {
      this.async.dispose();
    }
  }

  public componentDidMount(): void {
    this.loadPages(this.state.activeValues);
  }

  /**
   * Renders the SPListpicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    const label: string = this.props.label || "News Selector";
    return (
      <div>
        <Header title={label} />
        <table className={styles.termFieldTable}>
          <tbody>
            <tr>
              <td colSpan={2}>
                <Label>Select News Channel</Label>
              </td>
            </tr>
            <tr>
              <td>
                <TermPicker
                  context={this.props.context}
                  newsSelectorHostProps={this.props}
                  disabled={this.props.disabled}
                  value={this.state.activeValues.NewsChannel}
                  onChanged={this.termsFromPickerChanged}
                  allowMultipleSelections={this.props.allowMultipleSelections}
                  isTermSetSelectable={this.props.isTermSetSelectable}
                  disabledTermIds={this.props.disabledTermIds}
                  termsService={this.termsService}
                  resolveDelay={this.props.resolveDelay === undefined ? 500 : this.props.resolveDelay} // in future this can be bubbled upper to the settings
                />
              </td>
              <td className={styles.termFieldRow}>
                <IconButton disabled={this.props.disabled} iconProps={{ iconName: "Tag" }} onClick={this.onOpenPanel} />
              </td>
            </tr>
            <tr>
              <td colSpan={2}>
                <ChoiceGroup
                  selectedKey={this.state.activeValues.ActiveDisplayMode.toString()}
                  className={styles.displayModeChoice}
                  options={[
                    {
                      key: ActiveDisplayModeType.Latest.toString(),
                      text: "Latest (most recent published) article"
                    },
                    {
                      key: ActiveDisplayModeType.Specific.toString(),
                      text: "Specific article",
                      onRenderField: (props, render) => {
                        return (
                          <div>
                            {render!(props)}
                            <Dropdown
                              className={styles.pageSelectorDdl}
                              selectedKey={this.state.activeValues.ArticleId.toString()}
                              onRenderTitle={this.onRenderDropDownTitle}
                              options={this.state.pageDropDownOptions}
                              required={this.state.activeValues.ActiveDisplayMode === ActiveDisplayModeType.Specific}
                              disabled={this.state.activeValues.ActiveDisplayMode === ActiveDisplayModeType.Latest}
                              onChanged={this.onArticleDropDownChanged}
                            />
                          </div>
                        );
                      }
                    }
                  ]}
                  onChange={this.onDisplayModeChange}
                  label="Article Display Mode"
                  required={true}
                />
              </td>
            </tr>
          </tbody>
        </table>

        <FieldErrorMessage errorMessage={this.state.errorMessage} />

        <Panel
          isOpen={this.state.openPanel}
          hasCloseButton={true}
          onDismiss={this.onClosePanel}
          isLightDismiss={true}
          type={PanelType.medium}
          headerText={this.props.panelTitle}
          onRenderFooterContent={() => {
            return (
              <div className={styles.actions}>
                <PrimaryButton iconProps={{ iconName: "Save" }} text="Save" value="Save" onClick={this.onSave} />

                <DefaultButton iconProps={{ iconName: "Cancel" }} text="Cancel" value="Cancel" onClick={this.onClosePanel} />
              </div>
            );
          }}>
          {
            /* Show spinner in the panel while retrieving terms */
            this.state.termStoreLoaded === false ? <Spinner type={SpinnerType.normal} /> : ""
          }
          {
            /* Once the state is loaded, start rendering the term store, group, term sets */
            this.state.termStoreLoaded === true ? this.state.termStores.map((termStore: ITermStore, index: number) => {
              return (
                <div key={termStore.Id}>
                  {
                    !this.props.hideTermStoreName ? <h3>{termStore.Name}</h3> : null
                  }
                  {
                    termStore.Groups && termStore.Groups._Child_Items_ && termStore.Groups._Child_Items_.map((group) => {
                      return <TermGroup key={group.Id}
                        group={group}
                        termstore={termStore.Id}
                        termsService={this.termsService}
                        activeNodes={this.state.activeValues.NewsChannel}
                        changedCallback={this.termsChanged}
                        multiSelection={this.props.allowMultipleSelections}
                        isTermSetSelectable={this.props.isTermSetSelectable}
                        disabledTermIds={this.props.disabledTermIds} />;
                    })
                  }
                </div>
              );
            }) : ""
          }
        </Panel>
      </div>
    );
  }

  private onRenderDropDownTitle(options: IDropdownOption[]): JSX.Element {
    const option = options[0];
    return (
      <div>
        {!this.state.pagesLoaded && <Spinner type={SpinnerType.normal} />}
        {this.state.pagesLoaded && <span>{option.text}</span>}
      </div>
    );
  }

}
