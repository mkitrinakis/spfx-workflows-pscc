import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

import {
  FormConfigParser,
  FormRenderer,
  FormConfigLoader,
  ListItemService,
  FormValues,
  IFormConfig
} from 'workflows-core';

const LOG_SOURCE: string = 'FormFormCustomizer';

export interface IFormFormCustomizerProperties {
  configFileUrl: string;
}

export default class FormFormCustomizer
  extends BaseFormCustomizer<IFormFormCustomizerProperties> {

  private formRenderer: FormRenderer | undefined;
  private listItemService: ListItemService | undefined;
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
    const configLoader = new FormConfigLoader(this.context);
    this.isReadOnly = this.displayMode === FormDisplayMode.Display;

    try {
      const xmlText = await configLoader.loadXml(this.properties.configFileUrl);
      this.formConfig = FormConfigParser.parse(xmlText);
      const fieldNames = this.formConfig.fields.map((field) => field.internalName);
      this.initialValues = await this.listItemService.loadFieldValues(fieldNames, this.displayMode);
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
      onSave: async (values: FormValues) => this.handleSave(this.listItemService!, values)
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
    const savedValues = await listItemService.saveFieldValues(values, this.displayMode);
    this.formSaved();
    return savedValues;
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
