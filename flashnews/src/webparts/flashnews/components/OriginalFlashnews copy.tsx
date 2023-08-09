import * as React from 'react';
import Marquee from "react-fast-marquee";
import styles from './Flashnews.module.scss';
import { IFlashnewsProps } from './IFlashnewsProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { Item } from '@pnp/sp/items';
import { sp } from '@pnp/sp';
import { SPPermission } from '@microsoft/sp-page-context';
import { DisplayMode } from '@microsoft/sp-core-library';
import Popup from 'reactjs-popup';
import * as $ from 'jquery';
//import * jQuery from "jquery";
//import * as jQuery from "jquery";
require("jquery");

// const { JSDOM } = require( "jsdom" );
// const { window } = new JSDOM();
// const { document } = (new JSDOM('')).window;
// var $ = jQuery = require('jquery')(window);
//require("jquery")(window);


var initStage : boolean =true;

export interface IFlashnewsState{
  itemstore:any[];
}


export default class Flashnews extends React.Component<IFlashnewsProps, IFlashnewsState> {

  public constructor(props:IFlashnewsProps,state:IFlashnewsState){
    super(props);
   // SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js');
   // this.getColumnData = this.getColumnData.bind(this);
    this.state={itemstore:[]};
    
  }

public async componentDidMount(){
  await this.getColumnData();
}


public async  getColumnData(){
  debugger;
  const today = new Date();
  var filterString = this.props.FilterBy+` ge datetime'${today.toISOString()}'`;
  console.log(filterString);
  //const items: any[] = await sp.web.lists.getById(this.props.list)
  console.log("NOTHING IS IMPOSSIBLE");
  this.HideWebPart();
  const items:any[] = await sp.web.lists.getById(this.props.list).items.filter(filterString).getAll();
  if(items.length > 0)
  {
    console.log("greater than 0 is executing");
    initStage=false;
    this.setState({itemstore:items});
  }
  else
  {
    console.log("0 length is executing");
    initStage=false;
    this.setState({itemstore:items});
  }
  //this.setState({itemstore:items});
  //console.log(this.props.FilterValue);
}

public webpartHide(){
 // $("#myWebPartID").css('display','none');
  //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parent('div').parent('div').css('display','none');
  console.log("inside timer function");
  var shami = $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]');
  console.log($('div[data-viewport-id^="WebPart.FlashnewsWebPart"]'));
  console.log(shami[0]  +"only webpart div");
  var kashi = $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]') .parent('div').parent('div');
  console.log(kashi + "webpart with parent div");
  $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]') .parent('div').parent('div').hide();
  $(kashi.css('display','none'));
  return null;
}

// //  "jquery":"node_modules/jquery/dist/jquery.min.js",
public HideWebPart(){
  debugger;
  //alert("test");
  console.log("inside hide webpart");
  //$('div[id^="workbenchComman"]').hide();
  var s = $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parent().parent();
  //var parentDiv = document.getElementById("idFlashNewsWP").parentElement;
  //console.log(parentDiv); 

 // $("#idFlashNewsWP").hide();
  //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').closest('div').css('display','none');
  //console.log(s);
  //$(s).hide();
  //$(s).hasClass("hideme");
  //$(s).css('display','none');
  //$(s.prevAll('div'));
  //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').hi;

  //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]') .parent('div').parent('div').hide();
  
  console.log("success");

  //setTimeout(() => {this.webpartHide();}, 5000);
  //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parentElement('div').hidden=true;
  //$(flashWP.parentNode.style.display='none');
 // jQuery('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parent('div').closest('div').hidden=true;
 //var foundWebPartID = $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]'); //.parentNode.parentNode;
 //console.log(foundWebPartID);
// console.log(foundWebPartID.iindexOf('WebPart.FlashnewsWebPart'));

 //if(foundWebPartID && foundWebPartIDindexOf('WebPart.FlashnewsWebPart') > 0) 
 //{
   //console.log("Iam inside of found webpart");
   // $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parentNode.parentNode.hidden=true;
 //}

 // setInterval(() => {
  //  console.log("how are you");
    //$('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parentElement('div').hidden=true;
   // var flashWP = $('div[data-viewport-id^="WebPart.FlashnewsWebPart"]');
 // $(flashWP.parentNode.style.display='none');
  //jQuery('div[data-viewport-id^="WebPart.FlashnewsWebPart"]').parent('div').hidden=true;
//}, 17000);
 return null;
}


public render(): React.ReactElement<IFlashnewsProps> {
  return ( 
    <section id="idFlashNewsWP">
    {(this.state.itemstore.length > 0 && initStage != true)? (
      <div id="myWebPartID"  className={ styles.flashnews }>
          <div className={ styles.container }>
            <div className={ styles.row }><h3 style={{width:'174px',paddingLeft:'10px'}}>{this.props.Title}</h3>
            <Marquee play={true} direction={this.props.direction} speed={this.props.speed} pauseOnHover={true} gradient={false} style={{"background-color":this.props.bgcolor,"color":this.props.fgcolor,"font-family":this.props.fontname,"font-size":this.props.fontsize,"height":this.props.height,"width":this.props.width}} >
                {(this.state.itemstore.map((item,index) => (
                        <div>
                          <Popup trigger={<div>{item["Title"]}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*** </div>} position="right left">
                            <div style={{maxWidth:'300px',backgroundColor:this.props.descbgcolor,color:this.props.descfgcolor,fontSize:this.props.descfontsize,fontFamily:this.props.descfontname}}>{item["Description"]}</div>
                          </Popup>
                        </div>
                          )))} 
            </Marquee>
          </div></div></div>)
     : (<div><button style={{height:'60px',width:'200px'}} onClick={this.webpartHide()} ></button>
      {this.HideWebPart()}
      </div>)}
      </section>
          );
    } 
}




// console.log(this.state.itemstore.length);
  
//   /*if(this.state.itemstore.length === 0){
//    // console.log("inside render");
//     //this.HideWebPart();
//   } */
//   console.log(initStage);
//   {

/* private getsingleItem(columndata){
  var i=0;
  console.log(this.state.itemstore[i]["Title"]);
  {this.state.itemstore[i]["Title"] !=null ? <div><h1>iam here</h1></div>:<div><h1>I am out</h1></div>}
} */

/* private getsingleItem(columndata){
  console.log(this.state.itemstore.length);
  {this.state.itemstore.length > 0 ?
    <div>
  console.log("TTT");
  return(
   <div>
      <Ticker>
        {()=> <><h1>{columndata}</h1><img src="www.my-image-source.com/" alt=""/></> }
    </Ticker> 
    </div>
  ) </div>: ""}
} */

  // public mytestfunction()
  // {
  //   return(
  //     <div>
  //      <h1>"inside mytest functionTESSssssssssss"</h1>
  //     </div>
  //   )
  // }



/*
//   let today;// = new Date();
//   const searchtext = this.props.FilterValue
//   if(searchtext.search("today")){
//      today = new Date();
//      console.log("indis");
//   }
//   else
//   {
//     today=new Date(this.props.FilterValue);
//   }
//  console.log("not inite"+today)
  // get all the items from a list
  //const items: [] = await sp.web.lists.getById(this.props.list).items();
//.select(this.props.column,this.props.FilterBy)
  // const fieldInfo = await sp.web.lists.getById(this.props.list).fields.getByInternalNameOrTitle(this.props.FilterBy)();
  // console.log(fieldInfo.TypeDisplayName)
  // console.log(this.props.FilterBy);

  //ilter=(Created ge datetime'2019-09-13T00:00:00Z')and (Created le datetime'2019-09-14T00:00:00Z')
  //var filterString = this.props.FilterBy+' ' +this.props.condition+` datetime'${today.toISOString()}'`; //"Expires ge datetime '"+today.toISOString()+''
  */




  /* {(this.state.itemstore.map((item,index) => (
  <div>
   {this.getsingleItem("SIVA")}
   <h1>"TESSssssssssss"</h1>
  </div>
)))} 
<div className={ styles.flashnews }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>  */