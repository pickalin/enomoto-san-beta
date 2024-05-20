import React from 'react';
import ReactDOM from 'react-dom/client'
import './index.css'

// import ApiCalendar from "react-google-calendar-api";

//const ExcelJS = require('exceljs');
import * as ExcelJS from 'exceljs';
import { format, addBusinessDays } from 'date-fns';
// import * as DateFns from 'date-fns';

interface IProps {
  history:string[],
}
interface eventHeader {
  TITLE: string | null, 
  PROMOTER: string | null
}
interface salesPhase {
  MANAGEMENT_ID: ExcelJS.CellValue | null,
  SALES_CATEGORY: ExcelJS.CellValue | null,
  CLOSINGDAY: string | null,
  CLOSINGDAY_OPTION: string | null,
  NUMBER_DELIVERYDAY: string | null,
  NUMBER_DELIVERYDAY_OPTION: string | null,
  SEAT_DELIVERYDAY: string | null,
  SEAT_DELIVERYDAY_OPTION: string | null,
  TICKETING_STARTDAY: string | null,
  SWITCHINGDAY: string | null,
  RECEPTION_REPORTINGDAY: string | null,
  RECEPTION_REPORTINGDAY_OPTION: string | null,
  RAFFLINGDAY: string | null,
  REMIND_EXCHTICKET: string | null,
  SALES_STARTDAY: string | null
}
interface IState {
  client: any,
  accessToken: string,
  eventHeader: eventHeader | null,
  salesPhase1: salesPhase | null,
  salesPhase2: salesPhase | null,
  salesPhase3: salesPhase | null,
  salesPhase4: salesPhase | null,
  calendarContents: calContent[]
}
interface calContent {
  summary: string, date: string, calID: string, trColor: React.CSSProperties
}


function toAllDayString (value: ExcelJS.CellValue) :string | null{
  if (value){
    let cellvalue = String(value);
    let date = new Date(cellvalue);
    //Invalid Dateを判定
    if (!Number.isNaN(date.getTime())){
      return date.toLocaleDateString('sv-SE');
    }
    return null;
  } else {
    return null;
  }
}

function optionCheck (value: ExcelJS.CellValue) :string{
  if (value) {
    return String(value) + "！";
  } else {
    return "";
  }
}

const CLIENT_ID = "290369033762-558up7i9efu6g13f8ub5871k0h2l9lf3.apps.googleusercontent.com"
const SCOPE = 'https://www.googleapis.com/auth/calendar';

const calendarIDs = {
  brown : "confetti-web.com_bkhpcvteat0f4anue04ff7cqo8@group.calendar.google.com",
  blue : "confetti-web.com_28u4a67shama6e432sg9mhfrso@group.calendar.google.com",
  red : "confetti-web.com_kodu75u5tl0ngbul56e2pch5c4@group.calendar.google.com",
  yellow : "confetti-web.com_6qtt9cpih9dbmqoe7fo5bsvuh4@group.calendar.google.com",
  yellow_green : "hyouken@confetti-web.com",
  purple : "confetti-web.com_h6pui18d6nt9q4gjv71bmchu84@group.calendar.google.com"
}

// const startEventID = 10001;
// let eventID = startEventID;

class App extends React.Component<IProps, IState> {
  constructor(props: any) {
    super(props);
    this.state = { 
      client: null, 
      accessToken: '', 
      eventHeader: {
        TITLE: null, 
        PROMOTER: null
      },
      salesPhase1: {
        MANAGEMENT_ID: null,
        SALES_CATEGORY: null,
        CLOSINGDAY: null,
        CLOSINGDAY_OPTION: "",
        NUMBER_DELIVERYDAY: null,
        NUMBER_DELIVERYDAY_OPTION: "",
        SEAT_DELIVERYDAY: null,
        SEAT_DELIVERYDAY_OPTION: "",
        TICKETING_STARTDAY: null,
        SWITCHINGDAY: null,
        RECEPTION_REPORTINGDAY: null,
        RECEPTION_REPORTINGDAY_OPTION: "",
        RAFFLINGDAY: null,
        REMIND_EXCHTICKET: null,
        SALES_STARTDAY: null
      },
      salesPhase2: {
        MANAGEMENT_ID: null,
        SALES_CATEGORY: null,
        CLOSINGDAY: null,
        CLOSINGDAY_OPTION: "",
        NUMBER_DELIVERYDAY: null,
        NUMBER_DELIVERYDAY_OPTION: "",
        SEAT_DELIVERYDAY: null,
        SEAT_DELIVERYDAY_OPTION: "",
        TICKETING_STARTDAY: null,
        SWITCHINGDAY: null,
        RECEPTION_REPORTINGDAY: null,
        RECEPTION_REPORTINGDAY_OPTION: "",
        RAFFLINGDAY: null,
        REMIND_EXCHTICKET: null,
        SALES_STARTDAY: null
      },
      salesPhase3: {
        MANAGEMENT_ID: null,
        SALES_CATEGORY: null,
        CLOSINGDAY: null,
        CLOSINGDAY_OPTION: "",
        NUMBER_DELIVERYDAY: null,
        NUMBER_DELIVERYDAY_OPTION: "",
        SEAT_DELIVERYDAY: null,
        SEAT_DELIVERYDAY_OPTION: "",
        TICKETING_STARTDAY: null,
        SWITCHINGDAY: null,
        RECEPTION_REPORTINGDAY: null,
        RECEPTION_REPORTINGDAY_OPTION: "",
        RAFFLINGDAY: null,
        REMIND_EXCHTICKET: null,
        SALES_STARTDAY: null
      },
      salesPhase4: {
        MANAGEMENT_ID: null,
        SALES_CATEGORY: null,
        CLOSINGDAY: null,
        CLOSINGDAY_OPTION: "",
        NUMBER_DELIVERYDAY: null,
        NUMBER_DELIVERYDAY_OPTION: "",
        SEAT_DELIVERYDAY: null,
        SEAT_DELIVERYDAY_OPTION: "",
        TICKETING_STARTDAY: null,
        SWITCHINGDAY: null,
        RECEPTION_REPORTINGDAY: null,
        RECEPTION_REPORTINGDAY_OPTION: "",
        RAFFLINGDAY: null,
        REMIND_EXCHTICKET: null,
        SALES_STARTDAY: null
      },
      calendarContents : []
    }
  }  

  /** アクセストークン取得 */
  getToken = () => {
    if (!this.state.client) {
      console.log("clientがないです。")
      return;
    }
    this.state.client.requestAccessToken();
  };
  /** アクセストークン削除 */
  revokeToken = () => {
    (window as any).google.accounts.oauth2.revoke(this.state.accessToken, () => {
      console.log('access token revoked');
    });
  };

  componentDidMount(){
    const self = this;
    window.document.getElementById("file1")?.addEventListener("change", function() {
      console.log("#change");
      self.uploadFile();
    }); 
    const script = document.createElement('script');
    // script.src = "https://accounts.google.com/gsi/client";
    // script.onload = initClient;
    script.async = true;
    script.defer = true;
    document.body.appendChild(script);
    const initTokenClient = (window as any).google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPE,
      callback: (tokenResponse: any) => {
        console.log(tokenResponse)
        this.setState({accessToken: tokenResponse.access_token});
      },
    });
    this.setState({client: initTokenClient});
  }
  uploadFile(){
    const self = this;
    console.log("uploadFile");
    const files = document.querySelector<HTMLInputElement>('#file1')?.files;
    const fileObject = files![0]; 
    if (typeof fileObject === "undefined") {
      console.error("none, fileObject");
      return;
    }
    console.log(fileObject);
    const blobURL = window.URL.createObjectURL(fileObject);
    console.log(blobURL);
    const xhr = new XMLHttpRequest();
    xhr.onload = function() {
      const result = xhr.response; // ArrayBuffer
//      console.log(result);
      const data = new Uint8Array(result);
//      console.log(data);
      self.loadExcelData(data);
    }
    xhr.responseType = "arraybuffer";
    xhr.open("GET", blobURL);
    xhr.send();    
    console.log("start-upload");
  }
  async loadExcelData(data: any){
    try{
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      const worksheet = workbook.getWorksheet('チェック表紙');
      worksheet!.pageSetup = {orientation:'portrait'};
      const startRow = 41;
      const headingColumn = 1;
      const eHeader = this.state.eventHeader;
      const phase1 = this.state.salesPhase1;
      const phase2 = this.state.salesPhase2;
      const phase3 = this.state.salesPhase3;
      const phase4 = this.state.salesPhase4;
      const phases = [phase1, phase2, phase3, phase4];
      let rowH = worksheet?.getRow(1);
      let row = worksheet?.getRow(1);
      //headerステートにheaderを登録
      for (let i = 1; i < 50; i++) {
        rowH = worksheet!.getRow(i);
        if(rowH.getCell(headingColumn).value == "タイトル"){
          eHeader!.TITLE = String(rowH.getCell(2).result);
        }
        if(rowH.getCell(headingColumn).value == "請求元"){
          eHeader!.PROMOTER = String(rowH.getCell(2).result);
          console.log(eHeader);
          this.setState({eventHeader: eHeader});
        }
        //各仮phaseプロパティに公演IDと販売区分を格納
        if(rowH.getCell(headingColumn).value == "担当"){
          phases[0]!.MANAGEMENT_ID = rowH.getCell(9).value;
          phases[1]!.MANAGEMENT_ID = worksheet!.getRow(i+1).getCell(9).value;
          phases[2]!.MANAGEMENT_ID = rowH.getCell(13).value;
          phases[3]!.MANAGEMENT_ID = worksheet!.getRow(i+1).getCell(13).value;
          // console.log("公演コード：" + row.getCell(9).value);
        }
        if(rowH.getCell(headingColumn).value == "販売方法"){
          phases[0]!.SALES_CATEGORY = rowH.getCell(3).result || null //rowH.getCell(3).value;
          phases[1]!.SALES_CATEGORY = rowH.getCell(10).result || null //rowH.getCell(10).value;
          phases[2]!.SALES_CATEGORY = rowH.getCell(17).result || null //rowH.getCell(17).value;
          phases[3]!.SALES_CATEGORY = rowH.getCell(24).result || null //rowH.getCell(24).value;
        }
      }
      //各仮phaseプロパティにイベント日付を格納
      for (let p = 0; p <= 3; p++) {
        if (phases[p]?.SALES_CATEGORY) {
          for (let i = startRow; i < 70; i++) {
            row = worksheet!.getRow(i);
            if (row.getCell(headingColumn).value == "確定数報告or返券日") {
              phases[p]!.CLOSINGDAY = toAllDayString(row.getCell(p*7+2).value);
              phases[p]!.CLOSINGDAY_OPTION = optionCheck(row.getCell(p*7+4).value);
            }
            if (row.getCell(headingColumn).value == "席番納品日") {
              phases[p]!.SEAT_DELIVERYDAY = toAllDayString(row.getCell(p*7+2).value);
              phases[p]!.SEAT_DELIVERYDAY_OPTION = optionCheck(row.getCell(p*7+4).value);
            }
            if (row.getCell(headingColumn).value == "発券開始日時") {
              phases[p]!.TICKETING_STARTDAY = toAllDayString(row.getCell(p*7+2).value);
            }
            if (row.getCell(headingColumn).value == "受付数報告日") {
              console.log("受付数報告日" + row.getCell(p*7+2).value);
              phases[p]!.RECEPTION_REPORTINGDAY = toAllDayString(row.getCell(p*7+2).value);
              phases[p]!.RECEPTION_REPORTINGDAY_OPTION = optionCheck(row.getCell(p*7+4).value);
            }
            if (row.getCell(headingColumn).value == "抽選結果発表") {
              console.log("抽選結果発表" + row.getCell(p*7+2).value);
              phases[p]!.RAFFLINGDAY = toAllDayString(row.getCell(p*7+2).value);
            }
            if (row.getCell(headingColumn).value == "数納品日") {
              console.log("数納品日" + row.getCell(p*7+2).value);
              phases[p]!.NUMBER_DELIVERYDAY = toAllDayString(row.getCell(p*7+2).value);
              phases[p]!.NUMBER_DELIVERYDAY_OPTION = optionCheck(row.getCell(p*7+4).value);
            }
            if (row.getCell(headingColumn).value == "販売開始日時") {
              phases[p]!.SALES_STARTDAY = toAllDayString(row.getCell(p*7+2).value);
            }
            //SWITCHNGDAY
            //REMIND_EXCHTICKET
          } 
        }
      }
      console.log(phases[0]);
      this.setState({salesPhase1: phases[0]});
      this.setState({salesPhase2: phases[1]});
      this.setState({salesPhase3: phases[2]});
      this.setState({salesPhase4: phases[3]});

      const calContents :calContent[] = [];
      const allPhaseData = [
        this.state.salesPhase1, 
        this.state.salesPhase2, 
        this.state.salesPhase3,
        this.state.salesPhase4
      ];
      const salesCategorys = [
        JSON.stringify(this.state.salesPhase1?.SALES_CATEGORY),
        JSON.stringify(this.state.salesPhase2?.SALES_CATEGORY),
        JSON.stringify(this.state.salesPhase3?.SALES_CATEGORY),
        JSON.stringify(this.state.salesPhase4?.SALES_CATEGORY)
      ];
      const calHeader = String(this.state.eventHeader?.TITLE! + " " + this.state.eventHeader?.PROMOTER + " ");
      
      let returnDayCaption: string = "";
      
      for (let i = 0; i < allPhaseData.length; i++) {

        //受付数報告日をpush
        if (salesCategorys[i].indexOf('抽選') > -1 ) {
          calContents.push({
            summary: allPhaseData[i]?.RECEPTION_REPORTINGDAY_OPTION + "【受付数報告】" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: allPhaseData[i]?.RECEPTION_REPORTINGDAY!,
            calID: calendarIDs.blue,
            trColor: {background: 'blue'}
          })
        }
        //数納品日をpush
        if (allPhaseData[i]?.NUMBER_DELIVERYDAY != null) {
          calContents.push({
            summary: this.state.salesPhase1?.NUMBER_DELIVERYDAY_OPTION + "【数納品日】" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: allPhaseData[i]?.NUMBER_DELIVERYDAY!,
            calID: calendarIDs.red,
            trColor: {background: 'red'}
          })
        }
        //席納品日をpush
        if (allPhaseData[i]?.SEAT_DELIVERYDAY != null) {
          calContents.push({
            summary: allPhaseData[i]?.SEAT_DELIVERYDAY_OPTION + "【席納品日】" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: allPhaseData[i]?.SEAT_DELIVERYDAY!,
            calID: calendarIDs.red,
            trColor: {background: 'red'}
          })
        }
        //抽選日をpush
        if (allPhaseData[i]?.RAFFLINGDAY) {
          calContents.push({
            summary: "【抽選日】" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: allPhaseData[i]?.RAFFLINGDAY!,
            calID: calendarIDs.purple,
            trColor: {background: 'purple'}
          })
        }
        //返券日もしくは先行報告もしくは確定数報告をpush
        if (allPhaseData[i]?.CLOSINGDAY){
          if (salesCategorys[i].indexOf('先行') == -1) {
            console.log('先行じゃない')
            if (salesCategorys[i+1] == "当日引換券") {
              calContents.push({
                summary: "※在庫調整※" + allPhaseData[i]?.CLOSINGDAY_OPTION + " " + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
                date: allPhaseData[i]?.CLOSINGDAY!,
                calID: calendarIDs.brown,
                trColor: {background: 'brown'}
              })
            }
            if (salesCategorys[i+1] == 'null') {
              console.log('次のフェーズはない')
              calContents.push({
                summary: allPhaseData[i]?.CLOSINGDAY_OPTION + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
                date: allPhaseData[i]?.CLOSINGDAY!,
                calID: calendarIDs.brown,
                trColor: {background: 'brown'}
              })
            }
          } else {
            //次のフェーズがあるならば切替え日をpush
            if (salesCategorys[i+1]) {
              calContents.push({
                summary: "【" + salesCategorys[i] + "→" + salesCategorys[i+1] + "】" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID) + "→" + String(allPhaseData[i+1]?.MANAGEMENT_ID),
                date: format(addBusinessDays(new Date(allPhaseData[i+1]?.SALES_STARTDAY!), -1), 'yyyy-MM-dd'),
                calID: calendarIDs.blue,
                trColor: {background: 'blue'}
              })
            }
            if (salesCategorys[i].indexOf('抽選') > -1) {
              returnDayCaption = "【確定数報告】";
            } else {
              returnDayCaption = "【先行報告】"
            }
            calContents.push({
              summary: allPhaseData[i]?.CLOSINGDAY_OPTION + returnDayCaption + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
              date: allPhaseData[i]?.CLOSINGDAY!,
              calID: calendarIDs.blue,
              trColor: {background: 'blue'}
            })
          }
        }
        
        //発券開始日をpush
        if (allPhaseData[i]?.TICKETING_STARTDAY != null && salesCategorys[i].indexOf('数受') > -1) {
          calContents.push({
            summary: calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: allPhaseData[i]?.TICKETING_STARTDAY!,
            calID: calendarIDs.yellow,
            trColor: {background: 'yellow'}
          })
        }
        
        //引換券確認をpush
        if (salesCategorys[i] == "当日引換券" && i != 0) {
          calContents.push({
            summary: "※引換券確認※" + calHeader + String(allPhaseData[i]?.MANAGEMENT_ID),
            date: format(addBusinessDays(new Date(allPhaseData[i]?.SALES_STARTDAY!), -5), 'yyyy-MM-dd'),
            calID: calendarIDs.yellow_green,
            trColor: {background: 'yellowgreen'}
          })
        }
      }
      console.log(calContents);
      this.setState({calendarContents: calContents});

      alert("complete load data");
    } catch (e) {
      console.error(e);
      alert("Error, load data");
    }    
  }
  
  /** カレンダーの登録 */
  registerCalendar = () => {
    // let calendarID = "488f862ddf6b74175d724869ef70315fd84b89cb42bf2da8071d8be8f140dd09@group.calendar.google.com"
    // eventID = eventID+1
    // console.log(JSON.stringify(eventID))
  //     // 登録する内容

    for (let i = 0; i < this.state.calendarContents.length; i++){
      let registerResource = {
        kind: "calendar#event",
        // id: JSON.stringify(eventID),
        summary: this.state.calendarContents[i].summary,
        // iCalUID: "https://calendar.google.com/calendar/ical/488f862ddf6b74175d724869ef70315fd84b89cb42bf2da8071d8be8f140dd09%40group.calendar.google.com/private-b89a3aecfdb5dd8699424d99eddc54fd/basic.ics",
        start: {
          date: this.state.calendarContents[i].date,
        },
        end: {
          date: this.state.calendarContents[i].date,
        },
        }
      fetch(`https://www.googleapis.com/calendar/v3/calendars/${this.state.calendarContents[i].calID}/events`, {
        method: "POST",
        headers: {
          'Authorization': `Bearer ${this.state.accessToken}`
        },
        body: JSON.stringify(registerResource)
      }).then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`)
        }
        return response.json();
      }).then((data: any) => {
        console.log("登録した予定: ", data)
        alert("カレンダーに予定を登録しました！")
      }).catch((error) => {
        console.log("リクエスト中にエラーが発生しました: ", error)
      })
    }
  };

  render(){
    console.log(this.state);
    return(
    <div className="container">
      <h1>榎本さんβ</h1>
      <hr />
      File: <br />
      <input type="file" name="file1" id="file1" /><br />
      <hr className="my-1" />
      {/* TABLE */}
      {this.state.eventHeader?.TITLE != null &&
        <div>
          <h3>タイトル： {this.state.eventHeader?.TITLE}</h3>
          <table>
            <thead><tr><th>日付</th><th>登録する内容</th></tr></thead>
          <tbody>
            {
              this.state.calendarContents.map((value: any, index: number) => (
              <tr key={index} style={value.trColor}><td>{value.date}</td><td>{value.summary}</td></tr>
              ))
            }
          </tbody>
          </table>
        </div>
      }
      <div>
        <button onClick={() => this.registerCalendar()}>カレンダーの登録</button><br />
        <button onClick={this.getToken}>アクセストークン取得</button><br />
        <button onClick={this.revokeToken}>トークン破棄</button>
      </div>
    </div>
    )
  }
}


ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App history={[]}/>
  </React.StrictMode>,
)
