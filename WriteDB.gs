function writeDB() {
  //Locate sheet 
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow =sheet.getMaxRows();
  Logger.clear();
   
  //Timestamp
  var range=sheet.getRange(lastRow,1);
  var timestamp=range.getValue();
  timestamp= Utilities.formatDate(timestamp, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  
  //ID
  range=sheet.getRange(lastRow,2);
  var id=range.getValue();
  id=Utilities.formatString('%04d', id);  
  
  //ChartNo
  range=sheet.getRange(lastRow,3);
  var chartNo=range.getValue();
  chartNo=Utilities.formatString('%08d', chartNo);
  
  //此次問卷時程
  range=sheet.getRange(lastRow,4);
  var timePoint=range.getValue();  
  switch(timePoint){
    case "M0":
      timePoint="Mental_M0";
      break;
    case "M1":
      timePoint="Mental_M1";
      break;
    case "M2":
      timePoint="Mental_M2";
      break;
    case "M3":
      timePoint="Mental_M3";
      break;
    case "M4":
      timePoint="Mental_M4";
      break;
    case "M5":
      timePoint="Mental_M5";
      break;
    case "M6":
      timePoint="Mental_M6";
      break;
    case "M7":
      timePoint="Mental_M7";
      break;
    case "M8":
      timePoint="Mental_M8";
      break;
  }
  
  //DiagnosedYear
  range=sheet.getRange(lastRow,11);
  var diagnosedYear=range.getValue();
  
  //Height
  range=sheet.getRange(lastRow,8);
  var height=range.getValue();
  
  //Weight
  range=sheet.getRange(lastRow,7);
  var weight=range.getValue();  
  
  //備註
  range=sheet.getRange(lastRow,10);
  var remark=range.getValue();  
  
  //請問您最近一次的CD4檢驗日期
  range=sheet.getRange(lastRow,12);
  var latestCD4Date=range.getValue(); 
  if(!"".equals(latestCD4Date)){
    latestCD4Date= Utilities.formatDate(latestCD4Date, "GMT+8", "yyyy-MM-dd HH:mm:ss");  
  }
  
  //請問您最近一次的CD4檢驗數據
  range=sheet.getRange(lastRow,13);
  var latestCD4Count=range.getValue();   
  
  //請問您最近一次的病毒量檢驗日期
  range=sheet.getRange(lastRow,14);
  var latestVLDate=range.getValue();
  if(!"".equals(latestVLDate)){
    latestVLDate= Utilities.formatDate(latestVLDate, "GMT+8", "yyyy-MM-dd HH:mm:ss");      
  }
    
  //請問您最近一次的病毒量檢驗數據
  range=sheet.getRange(lastRow,15);
  var latestVLCount=range.getValue(); 
  
  //Occupation
  range=sheet.getRange(lastRow,16);
  var occupation=range.getValue();   
  
  //目前是否有家人或朋友知道您的疾病狀況？
  range=sheet.getRange(lastRow,17);
  var knownStatus=range.getValue();  
  
  //請問您是否曾經至精神科門診或診所就醫？
  range=sheet.getRange(lastRow,18);
  var psychyOPD=range.getValue();    
  switch(psychyOPD){
    case "是": 
      psychyOPD=1;
      break;
    case "否":
      psychyOPD=0;
      break;
    default:
      psychyOPD=null;
      break;
  }
  
  //您至精神科門診或診所就醫之原因為何？
  range=sheet.getRange(lastRow,19);
  var psychyOPDReason=range.getValue();    
  
  //是否曾經被精神科醫師診斷過精神相關疾病？
  range=sheet.getRange(lastRow,20);
  var diagnosedPsy=range.getValue();    
  switch(diagnosedPsy){
    case "是": 
      diagnosedPsy=1;
      break;
    case "否":
      diagnosedPsy=0;
      break;
    default:
      diagnosedPsy=null;
      break;
  }
  
  //請問您曾經被診斷過的精神相關疾病為何？
  range=sheet.getRange(lastRow,21);
  var diagnosedPsyType=range.getValue();  
  
  //您是否因上述精神疾病而固定於精神科門診或診所就醫？
  range=sheet.getRange(lastRow,22);
  var regularFUPsy=range.getValue();    
  switch(regularFUPsy){
    case "是": 
      regularFUPsy=1;
      break;
    case "否":
      regularFUPsy=0;
      break;
    default:
      regularFUPsy=null;
      break;
  }  
  
  //Patient's Email
  range=sheet.getRange(lastRow,23);
  var emailPT=range.getValue();  
  
  //CESD-Q1~4
  var cesd14=[];
  var count=0;
  for(var i=24;i<28;i++){
    range=sheet.getRange(lastRow,i);
    var depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd14[count]=0;
        count++;
        break;
      case "有時 (一周發生1-2天)":
        cesd14[count]=1;
        count++;
        break;
      case "經常 (一周發生3-4天)":
        cesd14[count]=2;       
        count++;
        break;
      case "總是 (一周發生5-7天)":
        cesd14[count]=3;   
        count++;
        break;
    }       
  }
  var cesd1=cesd14[0];
  var cesd2=cesd14[1];
  var cesd3=cesd14[2];
  var cesd4=cesd14[3];  
  
  //CESD-Q5
  var cesd5=0
  range=sheet.getRange(lastRow,28);
  depression=range.getValue(); 
  switch(depression){
    case "總是 (一周發生5-7天)":
      cesd5=0;
      break;
    case "經常 (一周發生3-4天)":
      cesd5=1;
      break;
    case "有時 (一周發生1-2天)":
      cesd5=2;       
      break;
    case "極少或從未發生 (一周發生<1天)":
      cesd5=3;   
      break;
  }    
  
  //CESD-Q6~7
  var cesd67=[];
  count=0;
  for(var i=29;i<31;i++){
    range=sheet.getRange(lastRow,i);
    depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd67[count]=0;
        count++;
        break;
      case "有時 (一周發生1-2天)":
        cesd67[count]=1;
        count++;
        break;
      case "經常 (一周發生3-4天)":
        cesd67[count]=2;       
        count++;
        break;
      case "總是 (一周發生5-7天)":
        cesd67[count]=3;   
        count++;
        break;
    }     
  }  
  var cesd6=cesd67[0];
  var cesd7=cesd67[1];  
  
  //CESD-Q8
  var cesd8=0;
  range=sheet.getRange(lastRow,31);
  depression=range.getValue(); 
  switch(depression){
    case "總是 (一周發生5-7天)":
      cesd8=0;
      break;
    case "經常 (一周發生3-4天)":
      cesd8=1;
      break;
    case "有時 (一周發生1-2天)":
      cesd8=2;       
      break;
    case "極少或從未發生 (一周發生<1天)":
      cesd8=3;   
      break;
  }     
 
  //CESD-Q9~10
  var cesd910=[];
  count=0;
  for(var i=32;i<34;i++){
    range=sheet.getRange(lastRow,i);
    depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd910[count]=0;
        count++;
        break;
      case "有時 (一周發生1-2天)":
        cesd910[count]=1;
        count++;
        break;
      case "經常 (一周發生3-4天)":
        cesd910[count]=2;       
        count++;
        break;
      case "總是 (一周發生5-7天)":
        cesd910[count]=3;   
        count++;
        break;
    }     
  } 
  var cesd9=cesd910[0];
  var cesd10=cesd910[1];
  
//PSQI
  //主觀睡眠品質Subjective sleep quality
  range=sheet.getRange(lastRow,49);
  var ssq=range.getValue();
  switch(ssq){
    case "非常好":
      ssq=0;
      break;
    case "好":
      ssq=1;
      break;
    case "不好":
      ssq=2;
      break;
    case "非常不好":
      ssq=3;
      break;
  }
  
  //延滯入睡時間Sleep onset latency
  //2. 過去一個月來，您在上床後通常多久才能入睡？
  range=sheet.getRange(lastRow,35);
  var sol1=range.getDisplayValue();
  var solArr=sol1.split(":")
  var solHr=solArr[0];
  var solMin=solArr[1];
  var solSec=solArr[2];
  solSec=solHr*60*60+solMin*60+solSec*1;
  if(solSec>3600){
    sol1=3;
  }else if(solSec>1800){
    sol1=2;
  }else if(solSec>900){
    sol1=1;
  }else{
    sol1=0;
  }
  
  //(1) 無法在 30 分鐘內入睡
  range=sheet.getRange(lastRow,38);
  var sol2=range.getDisplayValue();
  switch(sol2){
    case "過去一週從未發生":
      sol2=0;
      break;
    case "一週發生少於1次":
      sol2=1;
      break;
    case "一週發生1-2次":
      sol2=2;
      break;
    case "一週發生3次以上":
      sol2=3;
      break;
  }
  
  //延滯入睡時間總分
  var solSum=sol1+sol2;
  var sol; //Sleep onset latency
  if(solSum>=5){
    sol=3;
  }else if(solSum>=3){
    sol=2;
  }else if(solSum>=1){
    sol=1;
  }else{
    sol=0;
  }
  
  //總睡眠時間Sleep duration
  range=sheet.getRange(lastRow,37);
  var sd1=range.getDisplayValue();
  var sdArr=sd1.split(":");
  var sdHr=sdArr[0];
  var sdMin=sdArr[1];
  var sdSec=sdArr[2];
  sdSec=sdHr*60*60+sdMin*60+sdSec*1;
  var sd;
  if(sdSec>=25200){
    sd=0;
  }else if(sdSec>=21600){
    sd=1;
  }else if(sdSec>=18000){
    sd=2;
  }else{
    sd=3;
  }
  
  //習慣性睡眠效率Habitual sleep efficiency
  //AJ
  range=sheet.getRange(lastRow,36);
  var hse1=range.getValue();
    
  //AH
  range=sheet.getRange(lastRow,34);
  var hse2=range.getValue();
  
  //AJ-AH
  var hse12Difference = hse1.getTime() - hse2.getTime(); 
  if(hse12Difference<0){
    hse12Difference+=24*60*60*1000;
  }else if(hse12Difference==0){
    hse12Difference=12*60*60*1000;
  }
  //4. 在最近一個月內，您每天真正睡著的時間約有多久？
  range=sheet.getRange(lastRow,37);
  var sleepTime=range.getDisplayValue();
  var sleepTimeArr=[];
  var sleepTimeArr=sleepTime.split(":");
  var sleepTimeHr=sleepTimeArr[0];
  var sleepTimeMin=sleepTimeArr[1];
  var sleepTimeSec=sleepTimeHr*3600+sleepTimeMin*60;
  var sleepTimeSecMilliseconds=sleepTimeSec*1000;
  
  var hsePercent=sleepTimeSecMilliseconds/hse12Difference*100;
  var hse=0;
  if(hsePercent>=85){
    hse=0;
  }else if(hsePercent>=75){
    hse=1;
  }else if(hsePercent>=65){
    hse=2;
  }else{
    hse=3;
  } 
  
 //睡眠困擾sleep disturbance
    var sdSum=0;
    var sleepDisturbanceVar=0;
    var sleepDisturbance=0;
    
  range=sheet.getRange(lastRow,47);
  var sleepDisturbanceValue=range.getValue();
  if(sleepDisturbanceValue.equals("")){
    for(var i=39;i<47;i++){
        range=sheet.getRange(lastRow,i);
        var sleepDisturbanceValue=range.getValue();
        switch(sleepDisturbanceValue){
          case "過去一週從未發生":
            sleepDisturbanceVar=0;
            break;
          case "一週發生少於1次":
            sleepDisturbanceVar=1;
            break;
          case "一週發生1-2次":
            sleepDisturbanceVar=2;       
            break;
          case "一週發生3次以上":
            sleepDisturbanceVar=3;   
            break;
        }       
        sdSum +=sleepDisturbanceVar;
      }
      if(sdSum>=19){
        sleepDisturbance=3;
      }else if(sdSum>=10){
        sleepDisturbance=2;
      }else if(sdSum>=1){
        sleepDisturbance=1;
      }else{
        sleepDisturbance=0;
      }
  }else{
    for(var i=39;i<48;i++){
        range=sheet.getRange(lastRow,i);
        var sleepDisturbanceValue=range.getValue();
        switch(sleepDisturbanceValue){
          case "過去一週從未發生":
            sleepDisturbanceVar=0;
            break;
          case "一週發生少於1次":
            sleepDisturbanceVar=1;
            break;
          case "一週發生1-2次":
            sleepDisturbanceVar=2;       
            break;
          case "一週發生3次以上":
            sleepDisturbanceVar=3;   
            break;
          default:
            sleepDisturbanceVar=0;   
            break;
        }       
        sdSum +=sleepDisturbanceVar;
      }
      if(sdSum>=17){
        sleepDisturbance=3;
      }else if(sdSum>=9){
        sleepDisturbance=2;
      }else if(sdSum>=1){
        sleepDisturbance=1;
      }else{
        sleepDisturbance=0;
      }
  }
    
  //安眠藥的使用Use of sleeping medication
  range=sheet.getRange(lastRow,50);
  var useVar=range.getValue();
  var use=0;
  switch(useVar){
    case "從來沒有":
      use=0;
      break;
    case "一週少於1次":
      use=1;
      break;
    case "一週1-2次":
      use=2;
      break;
    case "一週3次以上":
      use=3;
      break;
  }
  
  //日間功能失調Daytime dysfunction
  range=sheet.getRange(lastRow,51);
  var daytimeValue=range.getValue();
  var daytimeVar1=0;
  switch(daytimeValue){
    case "從來沒有":
      daytimeVar1=0;
      break;
    case "一週少於1次":
      daytimeVar1=1;
      break;
    case "一週1-2次":
      daytimeVar1=2;
      break;
    case "一週3次以上":
      daytimeVar1=3;
      break;
  }
  
  range=sheet.getRange(lastRow,52);
  daytimeValue=range.getValue();
  var daytimeVar2=0;
    switch(daytimeValue){
    case "完全沒困難":
      daytimeVar2=0;
      break;
    case "沒困難":
      daytimeVar2=1;
      break;
    case "有困難":
      daytimeVar2=2;
      break;
    case "有很大的困難":
      daytimeVar2=3;
      break;
  }
  
  var daytimeSum=daytimeVar1+daytimeVar2;
  var daytimeDysfunction=0;
  if(daytimeSum>=5){
    daytimeDysfunction=3;
  }else if(daytimeSum>=3){
    daytimeDysfunction=2;
  }else if(daytimeSum>=1){
    daytimeDysfunction=1;
  }else{
    daytimeDysfunction=0;
  }
  
  var psqi=ssq+sol+sd+hse+sleepDisturbance+use+daytimeDysfunction;
  
  
  //DUDITE
  var dudite=[];
  count=0;
  for(var i=53;i<66;i++){
    range=sheet.getRange(lastRow,i);
    var drug=range.getValue();
    switch (drug){
      case "未曾使用":
        dudite[count]=0;
        count++;
        break;
      case "試過1次或以上":
        dudite[count]=1;
        count++;
        break;
      case "每月1次或更少":
        dudite[count]=2;
        count++;
        break;
      case "每月2-4次":
        dudite[count]=3;
        count++;
        break;
      case "每週2-3次":
        dudite[count]=4;
        count++;
        break;
      case "每週4次或以上":
        dudite[count]=5;
        count++;
        break;
    }
  }
  var d1=dudite[0];
  var d2=dudite[1];
  var d3=dudite[2];
  var d4=dudite[3];
  var d5=dudite[4];
  var d6=dudite[5];
  var d7=dudite[6];
  var d8=dudite[7];
  var d9=dudite[7];
  var d10=dudite[9]; 
  var d11=dudite[10]; 
  var d12=dudite[11]; 
  var d13=dudite[12];   
  
  var duditeP1=null;
  var duditeP2=null;
  var duditeP3=null;
  var duditeP4=null;
  var duditeP5=null;
  var duditeP6=null;
  var duditeP7=null;
  var duditeP8=null;
  var duditeP9=null;
  var duditeP10=null;
  var duditeP11=null;
  var duditeP12=null;
  var duditeP13=null;
  var duditeP14=null;
  var duditeP15=null;
  var duditeP16=null;
  var duditeP17=null;
  
  var duditeN1=null;
  var duditeN2=null;
  var duditeN3=null;
  var duditeN4=null;
  var duditeN5=null;
  var duditeN6=null;
  var duditeN7=null;
  var duditeN8=null;
  var duditeN9=null;
  var duditeN10=null;
  var duditeN11=null;
  var duditeN12=null;
  var duditeN13=null;
  var duditeN14=null;
  var duditeN15=null;
  var duditeN16=null;
  var duditeN17=null;
  
  var duditeT1=null;
  var duditeT2=null;
  var duditeT3=null;
  var duditeT4=null;
  var duditeT5=null;
  var duditeT6=null;
  var duditeT7=null;
  var duditeT8=null;
  var duditeT9=null;
  var duditeT10=null;
 
  if(d1>3||d2>3||d3>3||d4>3||d5>3||d6>3||d7>3||d8>3||d9>3||d10>3||d11>3||d12>3||d13>3){
    //DUDITE-P
    var duditeP=[];
    count=0;
    for(var i=67;i<85;i++){
      range=sheet.getRange(lastRow,i);
      var positive=range.getDisplayValue();
      switch(positive){
        case "沒有":
          duditeP[count]=0;
          count++;
          break;
        case "有一點點":
          duditeP[count]=1;
          count++;
          break;
        case "有一些":
          duditeP[count]=2;
          count++;
          break;
        case "有很多":
          duditeP[count]=3;
          count++;
          break;
        case "總是如此":
          duditeP[count]=4;
          count++;
          break;
        default:
          duditeP[count]=null;
          count++;
          break;
      }
    }
    duditeP1=duditeP[0];
    duditeP2=duditeP[1];
    duditeP3=duditeP[2];
    duditeP4=duditeP[3];
    duditeP5=duditeP[4];
    duditeP6=duditeP[5];
    duditeP7=duditeP[6];
    duditeP8=duditeP[7];
    duditeP9=duditeP[8];
    duditeP10=duditeP[9];
    duditeP11=duditeP[10];
    duditeP12=duditeP[11];
    duditeP13=duditeP[12];
    duditeP14=duditeP[13];
    duditeP15=duditeP[14];
    duditeP16=duditeP[15];
    duditeP17=duditeP[16];
    

  //DUDITE-N-Q1~Q4    
    var duditeN14=[];
    count=0;
    for(var i=84;i<88;i++){
      range=sheet.getRange(lastRow,i);
      var negative=range.getDisplayValue();
      switch(negative){
        case "沒有":
          duditeN14[count]=0;
          count++;
          break;
        case "每月少於1次":
          duditeN14[count]=1;
          count++;
          break;
        case "有一些":
          duditeN14[count]=2;
          count++;
          break;
        case "每個禮拜":
          duditeN14[count]=3;
          count++;
          break;
        case "每天或幾乎每天":
          duditeN14[count]=4;
          count++;
          break;
        default:
          duditeN14[count]=null;
          count++;
          break;
      }
    }
    duditeN1=duditeN14[0];
    duditeN2=duditeN14[1];
    duditeN3=duditeN14[2];
    duditeN4=duditeN14[3];
    
  //DUDITE-N-Q5~Q17
    var duditeN517=[];
    count=0;
    for(var i=88;i<101;i++){
      range=sheet.getRange(lastRow,i);
      negative=range.getDisplayValue();
      switch(negative){
        case "沒有":
          duditeN517[count]=0;
          count++;
          break;
        case "有一點點":
          duditeN517[count]=1;
          count++;
          break;
        case "有一些":
          duditeN517[count]=2;
          count++;
          break;
        case "有很多":
          duditeN517[count]=3;
          count++;
          break;
        case "總是如此":
          duditeN517[count]=4;
          count++;
          break;
        default:
          duditeN517[count]=null;
          count++;
          break;
      }
    }    
    duditeN5=duditeN517[0];
    duditeN6=duditeN517[1];
    duditeN7=duditeN517[2];
    duditeN8=duditeN517[3];
    duditeN9=duditeN517[4];
    duditeN10=duditeN517[5];
    duditeN11=duditeN517[6];
    duditeN12=duditeN517[7];
    duditeN13=duditeN517[8];
    duditeN14=duditeN517[9];
    duditeN15=duditeN517[10];
    duditeN16=duditeN517[11];
    duditeN17=duditeN517[12];
    
  //DUDITE-T
    var duditeT=[];
    count=0;
    for(var i=101;i<111;i++){
      range=sheet.getRange(lastRow,i);
      var treatment=range.getDisplayValue();
      switch(treatment){
        case "沒有":
          duditeT[count]=0;
          count++;
          break;
        case "有時候":
          duditeT[count]=1;
          count++;
          break;
        case "總是如此":
          duditeT[count]=2;
          count++;
          break;
        default:
          duditeT[count]=null;
          count++;
          break;
      }
    }    
    
    duditeT1=duditeT[0];
    duditeT2=duditeT[1];
    duditeT3=duditeT[2];
    duditeT4=duditeT[3];
    duditeT5=duditeT[4];
    duditeT6=duditeP[5];
    duditeT7=duditeT[6];
    duditeT8=duditeT[7];
    duditeT9=duditeT[8];
    duditeT10=duditeT[9];
  }
  
  //---Connect Azure SQL
  var dbName="yourDBaddress";
  var username="yourUsername";
  var password="yourPassword";
  var conn = Jdbc.getConnection(dbName, username, password);
  var insertTableSQL1 = "INSERT INTO Mental_Patient (ChartNo, ID, DiagnosedYear) VALUES (?,?,?)"; 
  var stmt1 = conn.prepareStatement(insertTableSQL1);
  stmt1.setString(1, chartNo);
  stmt1.setString(2, id);
  stmt1.setString(3, diagnosedYear);
  try{
  stmt1.execute();
  }catch(e){
    Logger.log("stmt1: "+chartNo+" "+e.message);
  }
    
    var insertTableSQL2 = "INSERT INTO "+ timePoint +"(Timestamp, ChartNo, Height, Weight, Remark, LatestCD4Date, LatestCD4Count, LatestVLDate, LatestVLCount, Occupation, KnownStatus, PsychyOPD, PsychyOPDReason, DiagnosedPsy, DiagnosedPsyType, RegularFUPsy, PTEmail, CESD1, CESD2, CESD3, CESD4, CESD5, CESD6, CESD7, CESD8, CESD9, CESD10, PSQI, DUDITEd1,DUDITEd2,DUDITEd3,DUDITEd4,DUDITEd5,DUDITEd6,DUDITEd7,DUDITEd8,DUDITEd9,DUDITEd10,DUDITEd11,DUDITEd12,DUDITEd13, DUDITEp1,DUDITEp2,DUDITEp3,DUDITEp4,DUDITEp5,DUDITEp6,DUDITEp7,DUDITEp8,DUDITEp9,DUDITEp10,DUDITEp11,DUDITEp12,DUDITEp13,DUDITEp14,DUDITEp15,DUDITEp16,DUDITEp17, DUDITEn1,DUDITEn2,DUDITEn3,DUDITEn4,DUDITEn5,DUDITEn6,DUDITEn7,DUDITEn8,DUDITEn9,DUDITEn10,DUDITEn11,DUDITEn12,DUDITEn13,DUDITEn14,DUDITEn15,DUDITEn16,DUDITEn17, DUDITEt1, DUDITEt2, DUDITEt3, DUDITEt4, DUDITEt5, DUDITEt6, DUDITEt7, DUDITEt8, DUDITEt9,DUDITEt10) VALUES (?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?)"; 
    var stmt2 = conn.prepareStatement(insertTableSQL2);
    stmt2.setString(1, timestamp);
    stmt2.setString(2, chartNo);
    stmt2.setString(3, height);
    stmt2.setString(4, weight);
    stmt2.setString(5, remark);
    stmt2.setString(6, latestCD4Date);
    stmt2.setString(7, latestCD4Count);
    stmt2.setString(8, latestVLDate);
    stmt2.setString(9, latestVLCount);
    stmt2.setString(10, occupation);
    stmt2.setString(11, knownStatus);
    stmt2.setString(12, psychyOPD);
    stmt2.setString(13, psychyOPDReason);
    stmt2.setString(14, diagnosedPsy);
    stmt2.setString(15, diagnosedPsyType);
    stmt2.setString(16, regularFUPsy);
    stmt2.setString(17, emailPT);  
    stmt2.setString(18, cesd1);
    stmt2.setString(19, cesd2);
    stmt2.setString(20, cesd3);
    stmt2.setString(21, cesd4);
    stmt2.setString(22, cesd5);
    stmt2.setString(23, cesd6);
    stmt2.setString(24, cesd7);
    stmt2.setString(25, cesd8);
    stmt2.setString(26, cesd9);
    stmt2.setString(27, cesd10);
    stmt2.setString(28, psqi);
    stmt2.setString(29, d1);
    stmt2.setString(30, d2);
    stmt2.setString(31, d3);
    stmt2.setString(32, d4);
    stmt2.setString(33, d5);
    stmt2.setString(34, d6);
    stmt2.setString(35, d7);
    stmt2.setString(36, d8);
    stmt2.setString(37, d9);
    stmt2.setString(38, d10);
    stmt2.setString(39, d11);
    stmt2.setString(40, d12);
    stmt2.setString(41, d13);
    
      stmt2.setString(42, duditeP1);
      stmt2.setString(43, duditeP2);
      stmt2.setString(44, duditeP3);
      stmt2.setString(45, duditeP4);
      stmt2.setString(46, duditeP5);
      stmt2.setString(47, duditeP6);
      stmt2.setString(48, duditeP7);
      stmt2.setString(49, duditeP8);
      stmt2.setString(50, duditeP9);
      stmt2.setString(51, duditeP10);
      stmt2.setString(52, duditeP11);
      stmt2.setString(53, duditeP12);
      stmt2.setString(54, duditeP13);
      stmt2.setString(55, duditeP14);
      stmt2.setString(56, duditeP15);
      stmt2.setString(57, duditeP16);
      stmt2.setString(58, duditeP17);
      stmt2.setString(59, duditeN1);
      stmt2.setString(60, duditeN2);
      stmt2.setString(61, duditeN3);
      stmt2.setString(62, duditeN4);
      stmt2.setString(63, duditeN5);
      stmt2.setString(64, duditeN6);
      stmt2.setString(65, duditeN7);
      stmt2.setString(66, duditeN8);
      stmt2.setString(67, duditeN9);
      stmt2.setString(68, duditeN10);
      stmt2.setString(69, duditeN11);
      stmt2.setString(70, duditeN12);
      stmt2.setString(71, duditeN13);
      stmt2.setString(72, duditeN14);
      stmt2.setString(73, duditeN15);
      stmt2.setString(74, duditeN16);
      stmt2.setString(75, duditeN17);
      stmt2.setString(76, duditeT1);
      stmt2.setString(77, duditeT2);
      stmt2.setString(78, duditeT3);
      stmt2.setString(79, duditeT4);
      stmt2.setString(80, duditeT5);
      stmt2.setString(81, duditeT6);
      stmt2.setString(82, duditeT7);
      stmt2.setString(83, duditeT8);
      stmt2.setString(84, duditeT9);
      stmt2.setString(85, duditeT10);
   
    
    
    try{
      stmt2.execute();
    }catch(e){
      MailApp.sendEmail("hank1992@gmail.com", "Mental Exception Report-"+ timePoint, "Message: " + chartNo+ " "+e.message + "\nFile: " + e.fileName + "\nLine: " + e.lineNumber);
      Logger.log("stmt2: "+chartNo+" "+e.message);
    }
  
  
  timePoint=timePoint.substring(7);
  var insertTableSQL3 = 
  "UPDATE dbo.Mental_Patient SET dbo.Mental_Patient."+timePoint+"Date=Mental_"+timePoint+".Timestamp FROM dbo.Mental_Patient INNER JOIN dbo.Mental_"+timePoint+" ON dbo.Mental_Patient.chartNo = dbo.Mental_"+timePoint+".chartNo\;"; 
  var stmt3 = conn.prepareStatement(insertTableSQL3);
  
  try{
      stmt3.execute();
    }catch(e){
      MailApp.sendEmail("hank1992@gmail.com", "Mental Exception Report-"+ timePoint, "Message: " + chartNo+ " "+e.message + "\nFile: " + e.fileName + "\nLine: " + e.lineNumber);
      Logger.log("stmt3: "+chartNo+" "+e.message);
    }
  
    conn.close();
  
}
