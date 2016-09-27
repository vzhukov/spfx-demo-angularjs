import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

// AngularJS
import * as angular from 'angular';
// Office UI Fabric
import 'ng-office-ui-fabric';
// ModuleLoader
import ModuleLoader from '@microsoft/sp-module-loader';
// EnvironmentType
import { EnvironmentType } from '@microsoft/sp-client-base';

import styles from './StuffClientWebPart.module.scss';
import * as strings from 'stuffClientWebPartStrings';
import { IStuffClientWebPartWebPartProps } from './IStuffClientWebPartWebPartProps';

export default class StuffClientWebPartWebPart extends BaseClientSideWebPart<IStuffClientWebPartWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);

    ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css');
    ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');
  }

  public render(): void {
    // Если метод render уже вызывался выходим из метода
    if(this.renderedOnce === true)
    {
      return;
    }
    angular
      .module('StuffApp', [
        'officeuifabric.core',
        'officeuifabric.components'])
      .controller('StuffController', ($scope: ng.IScope):void =>
      {
        // Паттерн фильтра
        ($scope as any).searchPattern = '';
        // Коллекция сотрудников
        ($scope as any).stuff = [];
        // ID списка
        var listId = '2B6B3A8F-6837-4141-A624-5373C2AC0816';
        // URL сайта
        var siteUrl = 'https://vz365.sharepoint.com/rusug'

        // Данные загружаем только если среда исполнения - SharePoint
        switch(this.context.environment.type)
        {
          case 2: // SharePoint
          case 3: // Classic server-rendered SharePoint
              // Запросы только через this.context.httpClient
              this.context.httpClient
                .get(siteUrl + "/_api/web/lists(guid'" + listId + "')/items")
                .then((response: Response) => {
                  response.json()
                  .then(data=>{
                    // Заполняем коллекцию сотрудников
                    ($scope as any).stuff = data.value;
                    if(!$scope.$$phase)
                    {
                      $scope.$apply();
                    }
                  });
                });
        }
      });
    // HTML разметка
    this.domElement.innerHTML = require('./templates/app.html');
    // Bootstrap
    angular.bootstrap(this.domElement, ['StuffApp']);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
