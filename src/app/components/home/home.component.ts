import { Component, OnDestroy } from '@angular/core';
import { getLocationOriginWithPath } from 'src/app/msal/msal-application.module';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
})
export class HomeComponent implements OnDestroy {

  showForceSignOut = false;

  public loginDialog!: Office.Dialog;

  constructor() {
  }

  ngOnDestroy(): void {
  }

  login() {
    const aadB2CRedirectUri = getLocationOriginWithPath('login');

    Office.context.ui.displayDialogAsync(aadB2CRedirectUri, { height: 86, width: 33 },
      (result) => {
        console.log(result.value);
        if (result.status === Office.AsyncResultStatus.Failed) {
          this.processLoginDialogEvent({ error: result.error.code });
        } else {
          this.loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, this.processLoginMessage);
          this.loginDialog.addEventHandler(Office.EventType.DialogEventReceived, this.processLoginDialogEvent);
        }
      });
  }

  processLoginMessage = (arg: any) => {
    console.log(arg);
  };

  processLoginDialogEvent = (arg: any) => {
    console.log(arg);

    this.loginDialog?.close();
    this.processDialogEvent(arg);
  };

  processDialogEvent = (arg: any) => {
    console.log(arg);

    switch (arg.error) {
      case 12002:
        break;
      case 12003:
        break;
      case 12006:
        // 12006 means that the user closed the dialog instead of waiting for it to close.
        // It is not known if the user completed the login or logout, so assume the user is
        // logged out and revert to the app's starting state. It does no harm for a user to
        // press the login button again even if the user is logged in.
        break;
      case 12007:
        break;
      default:
        break;
    }
  };
}
