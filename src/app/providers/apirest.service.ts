import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { productos } from './api_interface';


@Injectable({
  providedIn: 'root'
})
export class ApirestService {


  url="https://dummyjson.com/"

  constructor(
    private http : HttpClient
  ) { 

  }


  getProducts(){
    return this.http.get(this.url +'products')
  }

}
