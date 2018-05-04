import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { DataTableModule } from './data-table';
import { HttpModule } from '@angular/http';


import { AppComponent } from './app.component';


@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule, DataTableModule, HttpModule 
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
