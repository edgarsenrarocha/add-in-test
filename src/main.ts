/// <reference path='../node_modules/@types/office-js-preview/index.d.ts' />

import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

Office.initialize = () => {
  OfficeExtension.config.extendedErrorLogging = true;

  platformBrowserDynamic().bootstrapModule(AppModule)
    .catch(err => console.error(err));
};
