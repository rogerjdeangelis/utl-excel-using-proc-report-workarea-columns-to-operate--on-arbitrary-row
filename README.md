# utl-excel-using-proc-report-workarea-columns-to-operate--on-arbitrary-row
Excel using proc report workarea columns to operate on arbitrary rows 
    SAS-L: Excel using proc report workarea columns to operate on arbitrary rows                                                     
                                                                                                                                     
    When using 'ods excel' you cannot use noprint with this common problem?                                                          
    You need to set aside some 'working hiddon columns';                                                                             
                                                                                                                                     
    It appears that it is much easier to use grouping with compute blocks when input is sorted.                                      
    This may save you a lot of time if you tray to do complex compute block processing.                                              
                                                                                                                                     
    github                                                                                                                           
    https://tinyurl.com/y44am97p                                                                                                     
    https://github.com/rogerjdeangelis/utl-excel-using-proc-report-workarea-columns-to-operate--on-arbitrary-row/tree/master         
                                                                                                                                     
         Two Solution                                                                                                                
                                                                                                                                     
             1. Sort datastep report                                                                                                 
             2. Sort report report                                                                                                   
             3. sort datastep(add missing row) report                                                                                
    *_                   _                                                                                                           
    (_)_ __  _ __  _   _| |_                                                                                                         
    | | '_ \| '_ \| | | | __|                                                                                                        
    | | | | | |_) | |_| | |_                                                                                                         
    |_|_| |_| .__/ \__,_|\__|                                                                                                        
            |_|                                                                                                                      
    ;                                                                                                                                
                                                                                                                                     
    proc sort data=sashelp.class(obs=5 drop=name) out=clsSrt;                                                                        
      by sex;                                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
     40 obs WORK.CLSSRT total obs=5                                                                                                  
                                                                                                                                     
      SEX    AGE    HEIGHT    WEIGHT                                                                                                 
                                                                                                                                     
       F      13     56.5       84.0                                                                                                 
       F      13     65.3       98.0                                                                                                 
       M      14     69.0      112.5                                                                                                 
       F      14     62.8      102.5                                                                                                 
       M      14     63.5      102.5                                                                                                 
                                                                                                                                     
    *            _               _                                                                                                   
      ___  _   _| |_ _ __  _   _| |_                                                                                                 
     / _ \| | | | __| '_ \| | | | __|                                                                                                
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                 
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                
                    |_|                                                                                                              
    ;                                                                                                                                
                                                                                                                                     
                                                                                                                                     
    This simple EXCEL report is quite difficult to produce using just proc report?                                                   
    Hope I do not regret this.                                                                                                       
                                                                                                                                     
                                                                                                                                     
     d:/xls/class.xlsx                                                                                                               
        +---------------------------------------------------+                                                                        
        |    A       |     B      |    C       |    D       |                                                                        
        +---------------------------------------------------+                                                                        
     1  |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                                                        
        +------------+------------+------------+------------+                                                                        
     2  |    F       |    99      |    69      |  112.5     | ==> yellow highlight                                                   
        +------------+------------+------------+------------+                                                                        
     3  |    F       |    13      |    58      |  101.5     |                                                                        
        +------------+------------+------------+------------+                                                                        
     4  |                                                   |                                                                        
        +------------+------------+------------+------------+                                                                        
     5  |    M       |    13      |    58      |  101.5     | ==> yellow highlight                                                   
        +------------+------------+------------+------------+                                                                        
     6  |    M       |    13      |    58      |  101.5     |                                                                        
        +------------+------------+------------+------------+                                                                        
        |                                                   |                                                                        
        +---------------------------------------------------+                                                                        
                                                                                                                                     
        [CLASSREPORT]                                                                                                                
                                                                                                                                     
    *               _         _       _            _                                         _                                       
     ___  ___  _ __| |_    __| | __ _| |_ __ _ ___| |_ ___ _ __    _ __ ___ _ __   ___  _ __| |_                                     
    / __|/ _ \| '__| __|  / _` |/ _` | __/ _` / __| __/ _ \ '_ \  | '__/ _ \ '_ \ / _ \| '__| __|                                    
    \__ \ (_) | |  | |_  | (_| | (_| | || (_| \__ \ ||  __/ |_) | | | |  __/ |_) | (_) | |  | |_                                     
    |___/\___/|_|   \__|  \__,_|\__,_|\__\__,_|___/\__\___| .__/  |_|  \___| .__/ \___/|_|   \__|                                    
                                                          |_|              |_|                                                       
    ;                                                                                                                                
                                                                                                                                     
    proc sort data=sashelp.class(obs=5 drop=name) out=clsSrt;                                                                        
      by sex;                                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
    data  clsFix;                                                                                                                    
      retain wrkCol "";                                                                                                              
      set  clsSrt;                                                                                                                   
      by sex;                                                                                                                        
      if first.sex then wrkCol="x";                                                                                                  
      else wrkCol="";                                                                                                                
      wrkCol2=sex;                                                                                                                   
    run;quit;                                                                                                                        
                                                                                                                                     
    /*                                                                                                                               
    Up to 40 obs from CLSFIX total obs=5                                                                                             
                                                                                                                                     
           Working Columns                                                                                                           
          Do not show on report                                                                                                      
         ----------------------                                                                                                      
                                                                                                                                     
    Obs    WRKCOL    WRKCOL2  SEX    AGE    HEIGHT    WEIGHT                                                                         
                                                                                                                                     
     1       x          F      F      13     56.5       84.0                                                                         
     2                  F      F      13     65.3       98.0                                                                         
     3                  F      F      14     62.8      102.5                                                                         
     4       x          M      M      14     69.0      112.5                                                                         
     5                  M      M      14     63.5      102.5                                                                         
    */                                                                                                                               
                                                                                                                                     
                                                                                                                                     
    ods excel file="d:/xls/class.xlsx";                                                                                              
    ods excel options ( sheet_name = "CLASSREPORT");                                                                                 
                                                                                                                                     
    proc report data=clsFix(obs=5) nowd missing                                                                                      
         out=preRpt(rename=_break_=break);                                                                                           
    cols wrkCol2 sex  age height weight wrkCol ;                                                                                     
    DEFINE wrkCol / display "" style={cellwidth=0pt} format=$char1.;                                                                 
    DEFINE wrkCol2 / order noprint ;                                                                                                 
    DEFINE sex / display  ;                                                                                                          
    compute after wrkCol2;                                                                                                           
      val=" ";                                                                                                                       
      line val $2.;                                                                                                                  
    endcomp;                                                                                                                         
    compute wrkCol;                                                                                                                  
         if wrkCol="x" then call define(_row_, "Style", "Style = [background = yellow]");                                            
         else call define(_row_, "Style", "Style = [background = white]");                                                           
         wrkCol="";                                                                                                                  
    endcomp;                                                                                                                         
    run;quit;                                                                                                                        
                                                                                                                                     
    ods excel close;                                                                                                                 
                                                                                                                                     
    *               _     ____                              _                                                                        
     ___  ___  _ __| |_  |___ \   _ __ ___ _ __   ___  _ __| |_ ___                                                                  
    / __|/ _ \| '__| __|   __) | | '__/ _ \ '_ \ / _ \| '__| __/ __|                                                                 
    \__ \ (_) | |  | |_   / __/  | | |  __/ |_) | (_) | |  | |_\__ \                                                                 
    |___/\___/|_|   \__| |_____| |_|  \___| .__/ \___/|_|   \__|___/                                                                 
                                          |_|                                                                                        
    ;                                                                                                                                
                                                                                                                                     
    proc sort data=sashelp.class(obs=5 drop=name) out=clsSrt;                                                                        
      by sex;                                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
    proc report data=clsSrt nowd missing out=clsFix(rename=(_break_=wrkCol ));                                                       
       by sex;                                                                                                                       
       cols sex age height weight;                                                                                                   
       define sex /  order ;                                                                                                         
       break before sex/skip;                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
    ods excel file="d:/xls/class2.xlsx";                                                                                             
    ods excel options ( sheet_name = "CLASSREPORT");                                                                                 
                                                                                                                                     
    proc report data=clsFix nowd missing                                                                                             
         out=preRpt(rename=_break_=break);                                                                                           
    cols sex2 sex   age height weight wrkCol ;                                                                                       
    DEFINE wrkCol / display "" style={cellwidth=0pt} ;                                                                               
    DEFINE sex2 / order noprint ;                                                                                                    
    DEFINE sex / display  ;                                                                                                          
    compute after sex2;                                                                                                              
      val=" ";                                                                                                                       
      line val $2.;                                                                                                                  
    endcomp;                                                                                                                         
    compute wrkCol;                                                                                                                  
         if wrkCol="SEX" then call define(_row_, "Style", "Style = [background = yellow]");                                          
         else call define(_row_, "Style", "Style = [background = white]");                                                           
         wrkCol="";                                                                                                                  
    endcomp;                                                                                                                         
    run;quit;                                                                                                                        
                                                                                                                                     
    ods excel close;                                                                                                                 
                                                                                                                                     
                                                                                                                                     
                                                                                                                                     
    *               _         _       _            _                                                                                 
     ___  ___  _ __| |_    __| | __ _| |_ __ _ ___| |_ ___ _ __                                                                      
    / __|/ _ \| '__| __|  / _` |/ _` | __/ _` / __| __/ _ \ '_ \                                                                     
    \__ \ (_) | |  | |_  | (_| | (_| | || (_| \__ \ ||  __/ |_) |                                                                    
    |___/\___/|_|   \__|  \__,_|\__,_|\__\__,_|___/\__\___| .__/                                                                     
                                                          |_|                                                                        
      __         _     _             _       __                              _                                                       
     / /__ _  __| | __| |  _ __ ___ (_)___ __\ \   _ __ ___ _ __   ___  _ __| |_                                                     
    | |/ _` |/ _` |/ _` | | '_ ` _ \| / __/ __| | | '__/ _ \ '_ \ / _ \| '__| __|                                                    
    | | (_| | (_| | (_| | | | | | | | \__ \__ \ | | | |  __/ |_) | (_) | |  | |_                                                     
    | |\__,_|\__,_|\__,_| |_| |_| |_|_|___/___/ | |_|  \___| .__/ \___/|_|   \__|                                                    
     \_\                                     /_/           |_|                                                                       
    ;                                                                                                                                
                                                                                                                                     
                                                                                                                                     
    proc sort data=sashelp.class(obs=5 drop=name) out=clsSrt;                                                                        
      by sex;                                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
    options missing=' ';                                                                                                             
    data  clsFix;                                                                                                                    
      retain wrkCol "";                                                                                                              
      set  clsSrt;                                                                                                                   
      by sex;                                                                                                                        
      if first.sex then wrkCol="x";                                                                                                  
      output;                                                                                                                        
      if wrkCol="x" then call missing(of _all_);                                                                                     
      output;                                                                                                                        
    run;quit;                                                                                                                        
                                                                                                                                     
    /*                                                                                                                               
    Up to 40 obs from CLSFIX total obs=5                                                                                             
                                                                                                                                     
           Working Columns                                                                                                           
          Do not show on report                                                                                                      
         ----------------------                                                                                                      
                                                                                                                                     
    Obs    WRKCOL    WRKCOL2  SEX    AGE    HEIGHT    WEIGHT                                                                         
                                                                                                                                     
     1       x          F      F      13     56.5       84.0                                                                         
     2                  F      F      13     65.3       98.0                                                                         
     3                  F      F      14     62.8      102.5                                                                         
     4       x          M      M      14     69.0      112.5                                                                         
     5                  M      M      14     63.5      102.5                                                                         
    */                                                                                                                               
                                                                                                                                     
                                                                                                                                     
    ods excel file="d:/xls/class3.xlsx";                                                                                             
    ods excel options ( sheet_name = "CLASSREPORT");                                                                                 
                                                                                                                                     
    proc report data=clsFix nowd missing;                                                                                            
    cols  sex  age height weight wrkCol;                                                                                             
    DEFINE sex / display  ;                                                                                                          
    define wrkCol / noprint;                                                                                                         
    endcomp;                                                                                                                         
    compute wrkCol;                                                                                                                  
         if wrkCol="x" then call define(_row_, "Style", "Style = [background = yellow]");                                            
         wrkCol="";                                                                                                                  
    endcomp;                                                                                                                         
    run;quit;                                                                                                                        
                                                                                                                                     
    ods excel close;                                                                                                                 
                                                                                                                                     
