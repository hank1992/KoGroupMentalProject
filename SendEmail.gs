function SendEmails() {
  //Locate sheet 
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getMaxRows();
  
  //ChartNo
  var range=sheet.getRange(lastRow,3);
  var chartNo=range.getValue();
  chartNo=Utilities.formatString('%08d', chartNo);
    
  //Timestamp
  range=sheet.getRange(lastRow,1);
  var timestamp=range.getValue();
  timestamp= Utilities.formatDate(timestamp, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  
  //CM
  range=sheet.getRange(lastRow,5);
  var cm=range.getValue();
    switch (cm){
      case "白芸慧":
        cm="yhpai623@gmail.com";
        break;
      case "賴怡因":
        cm="yiin509@gmail.com";
        break;
      case "劉曉穎":
        cm="n038915@gmail.com";
        break;
      case "邱培仁":
        cm="d716gary@gmail.com";
        break;
      case "唐于雯":
        cm="winn618@gmail.com";
        break;
  }
  
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
  
  //CESD-Q1~4
  var cesdSUM=0;
  var cesd=0;
  for(var i=24;i<28;i++){
    range=sheet.getRange(lastRow,i);
    var depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd=0;
        break;
      case "有時 (一周發生1-2天)":
        cesd=1;
        break;
      case "經常 (一周發生3-4天)":
        cesd=2;       
        break;
      case "總是 (一周發生5-7天)":
        cesd=3;   
        break;
    }       
    cesdSUM +=cesd;
  }
  //CESD-Q5
    range=sheet.getRange(lastRow,28);
    depression=range.getValue(); 
    switch(depression){
      case "總是 (一周發生5-7天)":
        cesd=0;
        break;
      case "經常 (一周發生3-4天)":
        cesd=1;
        break;
      case "有時 (一周發生1-2天)":
        cesd=2;       
        break;
      case "極少或從未發生 (一周發生<1天)":
        cesd=3;   
        break;
    }    
    cesdSUM+=cesd;
  //CESD-Q6~7
  for(var i=29;i<31;i++){
    range=sheet.getRange(lastRow,i);
    depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd=0;
        break;
      case "有時 (一周發生1-2天)":
        cesd=1;
        break;
      case "經常 (一周發生3-4天)":
        cesd=2;       
        break;
      case "總是 (一周發生5-7天)":
        cesd=3;   
        break;
    }     
    cesdSUM +=cesd;
  }  
  //CESD-Q8
    range=sheet.getRange(lastRow,31);
    depression=range.getValue(); 
   switch(depression){
      case "總是 (一周發生5-7天)":
        cesd=0;
        break;
      case "經常 (一周發生3-4天)":
        cesd=1;
        break;
      case "有時 (一周發生1-2天)":
        cesd=2;       
        break;
      case "極少或從未發生 (一周發生<1天)":
        cesd=3;   
        break;
    }     
    cesdSUM +=cesd;  
  //CESD-Q9~10
  for(var i=32;i<34;i++){
    range=sheet.getRange(lastRow,i);
    depression=range.getValue();
    switch(depression){
      case "極少或從未發生 (一周發生<1天)":
        cesd=0;
        break;
      case "有時 (一周發生1-2天)":
        cesd=1;
        break;
      case "經常 (一周發生3-4天)":
        cesd=2;       
        break;
      case "總是 (一周發生5-7天)":
        cesd=3;   
        break;
    }     
    cesdSUM +=cesd;
  }      
  
  //DUDITE
  var duditeDCount=0;
  
  //大麻
  range=sheet.getRange(lastRow,53);
  var d1=range.getValue();
  switch (d1){
    case "未曾使用":
      d1=0;
      break;
    case "試過1次或以上":
      d1=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d1=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d1=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d1=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d1=5;
      duditeDCount++;
      break;
  }

  //安非他命
  range=sheet.getRange(lastRow,54);
  var d2=range.getValue();   
    switch (d2){
    case "未曾使用":
      d2=0;
      break;
    case "試過1次或以上":
      d2=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d2=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d2=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d2=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d2=5;
      duditeDCount++;
      break;
    }
  //古柯鹼
  range=sheet.getRange(lastRow,55);
  var d3=range.getValue();    
    switch (d3){
    case "未曾使用":
      d3=0;
      break;
    case "試過1次或以上":
      d3=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d3=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d3=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d3=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d3=5;
      duditeDCount++;
      break;
    }
  //海洛因
  range=sheet.getRange(lastRow,56);
  var d4=range.getValue();    
    switch (d4){
    case "未曾使用":
      d4=0;
      break;
    case "試過1次或以上":
      d4=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d4=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d4=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d4=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d4=5;
      duditeDCount++;
      break;
    }
  //K他命
  range=sheet.getRange(lastRow,57);
  var d5=range.getValue(); 
    switch (d5){
    case "未曾使用":
      d5=0;
      break;
    case "試過1次或以上":
      d5=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d5=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d5=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d5=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d5=5;
      duditeDCount++;
      break;
    }
  //搖頭丸
  range=sheet.getRange(lastRow,58);
  var d6=range.getValue();   
    switch (d6){
    case "未曾使用":
      d6=0;
      break;
    case "試過1次或以上":
      d6=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d6=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d6=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d6=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d6=5;
      duditeDCount++;
      break;
  }
  //其他幻覺劑
  range=sheet.getRange(lastRow,59);
  var d7=range.getValue(); 
    switch (d7){
    case "未曾使用":
      d7=0;
      break;
    case "試過1次或以上":
      d7=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d7=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d7=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d7=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d7=5;
      duditeDCount++;
      break;
    }
  //揮發劑或稀釋劑
  range=sheet.getRange(lastRow,60);
  var d8=range.getValue();
    switch (d8){
    case "未曾使用":
      d8=0;
      break;
    case "試過1次或以上":
      d8=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d8=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d8=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d8=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d8=5;
      duditeDCount++;
      break;
    }
  //GHB或其他藥物
  range=sheet.getRange(lastRow,61);
  var d9=range.getValue(); 
    switch (d9){
    case "未曾使用":
      d9=0;
      break;
    case "試過1次或以上":
      d9=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d9=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d9=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d9=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d9=5;
      duditeDCount++;
      break;
  }
  //安眠藥_鎮靜劑
  range=sheet.getRange(lastRow,62);
  var d10=range.getValue(); 
    switch (d10){
    case "未曾使用":
      d10=0;
      break;
    case "試過1次或以上":
      d10=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d10=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d10=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d10=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d10=5;
      duditeDCount++;
      break;
    }
  //止痛劑
  range=sheet.getRange(lastRow,63);
  var d11=range.getValue(); 
    switch (d11){
    case "未曾使用":
      d11=0;
      break;
    case "試過1次或以上":
      d11=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d11=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d11=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d11=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d11=5;
      duditeDCount++;
      break;
    }
  //酒
  range=sheet.getRange(lastRow,64);
  var d12=range.getValue(); 
    switch (d12){
    case "未曾使用":
      d12=0;
      break;
    case "試過1次或以上":
      d12=1;
      duditeDCount++;
      break;
    case "每月1次或更少":
      d12=2;
      duditeDCount++;
      break;
    case "每月2-4次":
      d12=3;
      duditeDCount++;
      break;
    case "每週2-3次":
      d12=4;
      duditeDCount++;
      break;
    case "每週4次或以上":
      d12=5;
      duditeDCount++;
      break;
    }
  //香菸_雪茄
  range=sheet.getRange(lastRow,65);
  var d13=range.getValue();
    switch (d13){
    case "未曾使用":
      d13=0;
      break;
    case "試過1次或以上":
      d13=1;
      break;
    case "每月1次或更少":
      d13=2;
      break;
    case "每月2-4次":
      d13=3;
      break;
    case "每週2-3次":
      d13=4;
      break;
    case "每週4次或以上":
      d13=5;
      break;
    }
  
  if(d1>3||d2>3||d3>3||d4>3||d5>3||d6>3||d7>3||d8>3||d9>3||d10>3||d11>3||d12>3||d13>3){
    //DUDITE-P
    for(var i=67;i<85;i++){
      range=sheet.getRange(lastRow,i);
      var positive=range.getValue();
      var duditeP;
      switch(positive){
        case "沒有":
          duditeP=0;
          break;
        case "有一點點":
          duditeP=1;
          break;
        case "有一些":
          duditeP=2;
          break;
        case "有很多":
          duditeP=3;
          break;
        case "總是如此":
          duditeP=4;
          break;
      }
    }
  //DUDITE-N-Q1~Q4
    for(var i=84;i<88;i++){
      range=sheet.getRange(lastRow,i);
      var negative=range.getValue();
      var duditeN;
      switch(negative){
        case "沒有":
          duditeN=0;
          break;
        case "每月少於1次":
          duditeN=1;
          break;
        case "每個月":
          duditeN=2;
          break;
        case "每個禮拜":
          duditeN=3;
          break;
        case "每天或幾乎每天":
          duditeN=4;
          break;
      }
    }
  //DUDITE-N-Q5~Q17
      for(var i=88;i<101;i++){
      range=sheet.getRange(lastRow,i);
      var negative=range.getValue();
      var duditeN;
      switch(negative){
        case "沒有":
          duditeN=0;
          break;
        case "有一點點":
          duditeN=1;
          break;
        case "有一些":
          duditeN=2;
          break;
        case "有很多":
          duditeN=3;
          break;
        case "總是如此":
          duditeN=4;
          break;
      }
    }
  //DUDITE-T1
      range=sheet.getRange(lastRow,101);
      var t1=range.getValue();
      switch(t1){
        case "沒有":
          t1=0;
          break;
        case "有時候":
          t1=1;
          break;
        case "總是如此":
          t1=2;
          break;
      }
  //DUDITE-T2
      range=sheet.getRange(lastRow,102);
      var t2=range.getValue();
      switch(t2){
        case "沒有":
          t2=0;
          break;
        case "有時候":
          t2=1;
          break;
        case "總是如此":
          t2=2;
          break;
      }
  //DUDITE-T3
      range=sheet.getRange(lastRow,103);
      var t3=range.getValue();
      switch(t3){
        case "沒有":
          t3=0;
          break;
        case "有時候":
          t3=1;
          break;
        case "總是如此":
          t3=2;
          break;
      }
  //DUDITE-T4
      range=sheet.getRange(lastRow,104);
      var t4=range.getValue();
      switch(t4){
        case "沒有":
          t4=0;
          break;
        case "有時候":
          t4=1;
          break;
        case "總是如此":
          t4=2;
          break;
      }
      //DUDITE-T5
      range=sheet.getRange(lastRow,105);
      var t5=range.getValue();
      switch(t5){
        case "沒有":
          t5=0;
          break;
        case "有時候":
          t5=1;
          break;
        case "總是如此":
          t5=2;
          break;
      }
      //DUDITE-T6
      range=sheet.getRange(lastRow,106);
      var t6=range.getValue();
      switch(t6){
        case "沒有":
          t6=0;
          break;
        case "有時候":
          t6=1;
          break;
        case "總是如此":
          t6=2;
          break;
      }
      //DUDITE-T7
      range=sheet.getRange(lastRow,107);
      var t7=range.getValue();
      switch(t7){
        case "沒有":
          t7=0;
          break;
        case "有時候":
          t7=1;
          break;
        case "總是如此":
          t7=2;
          break;
      }
      //DUDITE-T8
      range=sheet.getRange(lastRow,108);
      var t8=range.getValue();
      switch(t8){
        case "沒有":
          t8=0;
          break;
        case "有時候":
          t8=1;
          break;
        case "總是如此":
          t8=2;
          break;
      }
      //DUDITE-T9
      range=sheet.getRange(lastRow,109);
      var t9=range.getValue();
      switch(t9){
        case "沒有":
          t9=0;
          break;
        case "有時候":
          t9=1;
          break;
        case "總是如此":
          t9=2;
          break;
      }
      //DUDITE-T10
      range=sheet.getRange(lastRow,110);
      var t10=range.getValue();
      switch(t10){
        case "沒有":
          t10=0;
          break;
        case "有時候":
          t10=1;
          break;
        case "總是如此":
          t10=2;
          break;
      }
  }
   
  var messageCESD="(⊙０⊙) 看起來似乎您的心情不太好，建議與您的個管師討論\n";
  var messagePSQI="(⊙０⊙) 看起來似乎您睡眠狀況不太好，建議與您的個管師討論\n";
  var messageDUDITE="(⊙０⊙) 看起來您似乎有菸酒或用藥的問題，建議與您的個管師討論\n";
  var messageCM="";
  var messagePT="";
  
  if(cesdSUM>10||psqi>5||d1>2||d2>2||d3>2||d4>2||d5>2||d6>2||d7>2||d8>2||d9>2||d10>2||d11>2||d12>4||d13>4){
    
    if(cesdSUM>10){
      messageCM +=  messageCESD;
      messagePT +=  messageCESD;
    }
  
    if(psqi>5){
      messageCM +=  messagePSQI;
      messagePT += messagePSQI;
    }
    
    if(d1>2||d2>2||d3>2||d4>2||d5>2||d6>2||d7>2||d8>2||d9>2||d10>2||d11>2||d12>4||d13>4){
      messageCM +=  messageDUDITE;
      messagePT += messageDUDITE;
    }
    
    //Write Email to CM + Assistant
    var recipientCM=cm;
    var subjectCM =timestamp + "身心計畫轉介評估問卷回饋"+ chartNo;
    messageCM +="感謝您耐心填寫問卷，針對您填寫之內容分析如下~\n☆☆憂鬱量表(CES-D)："+cesdSUM+"分\n說明：總得分10分以上者，表示有憂鬱傾向，建議轉介身心科評估治療。\n☆☆睡眠品質量表(PSQI)："+psqi+"分\n說明:總得分範圍 0-21分，分數越高代表睡眠品質越不佳。\n☆☆成癮行為嚴重度量表(DUDITE)：\nDrug使用頻率如下\n"+"1. 大麻: "+d1+"\n2. 安非他命: "+d2+"\n3. 古柯鹼: "+d3+"\n4. 海洛因: "+d4+"\n5. K他命: "+d5+"\n6. 搖頭丸: "+d6+"\n7. 其他幻覺劑: "+d7+"\n8. 揮發劑或稀釋劑: "+d8+"\n9. GHB (G水) 或其他藥物: "+d9+"\n10. 安眠藥/鎮靜劑: "+d10+"\n11. 止痛劑: "+d11+"\n12. 酒: "+d12+"\n13. 香菸、雪茄: "+d13+"\n說明: 0分: 未曾使用, 1分: 試過1次或以上, 2分: 每月1次或更少, 3分: 每月2-4次, 4分: 每週2-3次, 5分: 每週4次或以上\n成大醫院 愛管閒事 關心您\n";
    MailApp.sendEmail(cm+", helperhiv@gmail.com, naiyingko@gmail.com", subjectCM, messageCM);
    //MailApp.sendEmail("hank1992@gmail.com", subjectCM, messageCM);
    
    //Write Email to Patient
    range=sheet.getRange(lastRow,23);
    var recipientPT=range.getValue();
    if(recipientPT.equals("")!=true){
      var subjectPT="問卷自動回饋";
      messagePT +="感謝您耐心填寫問卷，針對您填寫之內容分析如下~\n☆☆憂鬱量表(CES-D)："+cesdSUM+"分\n說明：總得分10分以上者，表示有憂鬱傾向，建議轉介身心科評估治療。\n☆☆睡眠品質量表(PSQI)："+psqi+"分\n說明:總得分範圍 0-21分，分數越高代表睡眠品質越不佳。\n☆☆曾使用藥物及酒精共"+duditeDCount+"種\n成大醫院 愛管閒事 關心您"
      MailApp.sendEmail(recipientPT, subjectPT, messagePT);
    }
    
  }else{
    //Write Email to Patient + Assistant
    range=sheet.getRange(lastRow,23);
    var recipientPT=range.getValue();
    if(recipientPT.equals("")!=true){
      var subjectPT="問卷自動回饋";
      messagePT="٩(●ᴗ●)۶ 太棒了!!您看起來無情緒及睡眠困擾\n感謝您耐心填寫問卷，針對您填寫之內容分析如下~\n☆☆憂鬱量表(CES-D)："+cesdSUM+"分\n說明：總得分10分以上者，表示有憂鬱傾向，建議轉介身心科評估治療。\n☆☆睡眠品質量表(PSQI)："+psqi+"分\n說明:總得分範圍 0-21分，分數越高代表睡眠品質越不佳。\n成大醫院 愛管閒事 關心您"
      MailApp.sendEmail(recipientPT+", helperhiv@gmail.com, naiyingko@gmail.com", subjectPT, messagePT);
    }else{
      //Write Email to Assistant  
      var subjectPT="問卷自動回饋";
      messagePT="٩(●ᴗ●)۶ 太棒了!!您看起來無情緒及睡眠困擾\n感謝您耐心填寫問卷，針對您填寫之內容分析如下~\n☆☆憂鬱量表(CES-D)："+cesdSUM+"分\n說明：總得分10分以上者，表示有憂鬱傾向，建議轉介身心科評估治療。\n☆☆睡眠品質量表(PSQI)："+psqi+"分\n說明:總得分範圍 0-21分，分數越高代表睡眠品質越不佳。\n成大醫院 愛管閒事 關心您"
      MailApp.sendEmail("helperhiv@gmail.com, naiyingko@gmail.com", subjectPT, messagePT);
    }
  }
}
