import { Component, OnInit } from '@angular/core';
import { app } from '@microsoft/teams-js';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  title = 'sample-teams-app';

  ngOnInit() {
    app.initialize()
    app.getContext().then((context: app.Context) => {
      console.log('context', context);
    }).catch((error: any) => {
      console.error('Error', error);
    });
  }

}
