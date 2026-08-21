import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

import {
  FormConfigParser,
  FormRenderer,
  FormConfigLoader,
  ListItemService,
  ActionExecutionService,
  FormDefaultValues,
  FormValues,
  IFormButtonConfig,
  IFormConfig
} from 'workflows-core';

import { registerLocalActions } from './actions/registerLocalActions';

const LOG_SOURCE: string = 'FormFormCustomizer';

export interface IFormFormCustomizerProperties {
  configFileUrl: string;
}

export default class FormFormCustomizer
  extends BaseFormCustomizer<IFormFormCustomizerProperties> {

  private formRenderer: FormRenderer | undefined;
  private listItemService: ListItemService | undefined;
  private actionExecutionService: ActionExecutionService | undefined;
  private formConfig: IFormConfig | undefined;
  private initialValues: FormValues = {};
  private loadError: string | undefined;
  private isReadOnly: boolean = false;

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FormFormCustomizer');

    if (!this.properties.configFileUrl) {
      this.loadError = 'Missing required property "configFileUrl".';
      this.render();
      return;
    }

    this.listItemService = new ListItemService(this.context);
    this.actionExecutionService = new ActionExecutionService();
    registerLocalActions();
    const configLoader = new FormConfigLoader(this.context);
    this.isReadOnly = this.displayMode === FormDisplayMode.Display;

    try {
      const xmlText = await configLoader.loadXml(this.properties.configFileUrl);
      this.formConfig = FormConfigParser.parse(xmlText);
      const fieldNames = this.formConfig.fields;
      this.initialValues = await this.listItemService.loadFieldValues(fieldNames, this.displayMode);

      if (this.displayMode === FormDisplayMode.New) {
        this.initialValues = FormDefaultValues.apply(this.formConfig.fields, this.initialValues);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error while rendering form.';
      Log.error(LOG_SOURCE, new Error(message));
      this.loadError = message;
    }

    this.render();
  }

  @override
  public render(): void {
    if (this.formRenderer) {
      this.formRenderer.dispose();
      this.formRenderer = undefined;
    }

    if (this.loadError) {
      this.renderError(this.loadError);
      return;
    }

    if (!this.formConfig || !this.listItemService) {
      this.renderLoading();
      return;
    }

    this.formRenderer = new FormRenderer({
      rootElement: this.domElement,
      config: this.formConfig,
      initialValues: this.initialValues,
      isReadOnly: this.isReadOnly,
      showSaveButton: !this.isReadOnly,
      formContext: this.context,
      onSave: async (values: FormValues) => this.handleSave(this.listItemService!, values),
      onActionButton: async (buttonConfig: IFormButtonConfig, values: FormValues) =>
        this.handleActionButton(buttonConfig, values)
    });

    this.formRenderer.render();
  }

  @override
  public onDispose(): void {
    if (this.formRenderer) {
      this.formRenderer.dispose();
    }
  }

  private async handleSave(
    listItemService: ListItemService,
    values: FormValues
  ): Promise<FormValues> {
    const saveResult = await listItemService.saveFieldValues(
      values,
      this.displayMode,
      this.formConfig!.fields
    );

    if (this.formConfig!.onSaveActions && this.formConfig!.onSaveActions.length > 0) {
      await this.actionExecutionService!.runOnSaveActions(this.formConfig!, {
        context: this.context,
        itemId: saveResult.itemId,
        formValues: saveResult.values,
        fields: this.formConfig!.fields,
        displayMode: this.displayMode
      });
    }

    this.formSaved();
    return saveResult.values;
  }

  private async handleActionButton(
    buttonConfig: IFormButtonConfig,
    values: FormValues
  ): Promise<void> {
    const itemId = this.displayMode === FormDisplayMode.New ? 0 : this.context.itemId;

    if (!itemId) {
      throw new Error('Save the item before running this action.');
    }

    await this.actionExecutionService!.runButtonAction(buttonConfig, {
      context: this.context,
      itemId: itemId,
      formValues: values,
      fields: this.formConfig!.fields,
      displayMode: this.displayMode
    });
  }

  private renderLoading(): void {
    this.domElement.innerHTML = '';
    this.domElement.classList.add('xml-form-root');

    const loading = document.createElement('div');
    loading.className = 'xml-form-loading';
    loading.textContent = 'Loading form...';
    this.domElement.appendChild(loading);
  }

  private renderError(message: string): void {
    this.domElement.innerHTML = '';
    this.domElement.classList.add('xml-form-root');

    const errorBox = document.createElement('div');
    errorBox.className = 'xml-form-error';
    errorBox.textContent = message;
    this.domElement.appendChild(errorBox);
  }
}
