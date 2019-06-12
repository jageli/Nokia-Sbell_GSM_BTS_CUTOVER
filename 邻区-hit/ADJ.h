string ERRORLINE;
integer  adj_file,inadj_file,adj3G_file,adjTD_file;
integer h_excel,new_lac_row;
string bsc_s,bcf_s,bts_s,topic,lac,ci;
string line,celllist[],laclist[],obtslist[],newet[],table[][]; 
integer cur_row,total_row;
string temp_buff;




//////define row name///
#define OLDBSC     1 
#define OLDLAC     2 
#define OLDCI      3 
#define OLDBCF     4 
#define OLDSEG     5 
#define OLDBTS     6 
#define OLDET      7 
#define NEWBSC     8 
#define NEWLAC     9 
#define NEWCI      10
#define NEWBCF     11
#define NEWSEG     12
#define NEWBTS     13
#define NEWET      14
#define NEWNSEI    15
#define NEWDAP     16
///end of define row name//
integer function ifbcfexist(string bscname_t,string bcfid_t,integer row_t)
integer i;
 	for (i = 1; i < row_t; i++)
 	    //stop();
		 if ((strtrim(bscname_t,1|2)==strtrim(cells(i,OLDBSC),1|2)) AND
		 	     (strtoint(bcfid_t,DEC)==strtoint(cells(i,OLDBCF),DEC))) 
		 			return 1;
		 endif
	endfor
	return 0;
endfunction

integer function ifsegexist(string bscname_t,string segid_t,integer row_t)
integer i;
 	for (i = 1; i < row_t; i++)
		 if ((strtrim(bscname_t,1|2)==strtrim(cells(i,OLDBSC),1|2)) AND
		  (strtoint(segid_t,DEC)==strtoint(cells(i,OLDSEG),DEC))) 
		 			return 1;
		 endif
	endfor
	return 0;
endfunction

string function mymid(string li,integer x, integer y)
string temp_str="";
if (li<>"")
strfetch(li,inttostr(x,DEC)+"-"+inttostr(x+y-1,DEC),temp_str);
return strtrim(strreplace(temp_str," ",""),1|2);
else
return "";
endif
endfunction

string function cells_ex(integer x,integer y)
  return strtrim(dderequest(h_excel,strprint("R%dC%d",x,y)),1|2);
endfunction

string function cells(integer x,integer y)
  return table[x][y];
endfunction

function readtotable()   
   cur_row=1;
	 total_row=1;
	 bsc_s=cells_ex(cur_row,1);
   while (bsc_s<>"")    
   //stop();
   table[cur_row][OLDBSC    ]=cells_ex(cur_row,OLDBSC    );
   table[cur_row][OLDBCF    ]=cells_ex(cur_row,OLDBCF    );
   table[cur_row][OLDSEG    ]=cells_ex(cur_row,OLDSEG    );
   table[cur_row][OLDBTS    ]=cells_ex(cur_row,OLDBTS    );
   table[cur_row][OLDLAC    ]=cells_ex(cur_row,OLDLAC    );
   table[cur_row][OLDCI     ]=cells_ex(cur_row,OLDCI     );
   table[cur_row][OLDET     ]=cells_ex(cur_row,OLDET     );
   table[cur_row][NEWBSC    ]=cells_ex(cur_row,NEWBSC    );
   table[cur_row][NEWBCF    ]=cells_ex(cur_row,NEWBCF    );
   table[cur_row][NEWSEG    ]=cells_ex(cur_row,NEWSEG    );
   table[cur_row][NEWBTS    ]=cells_ex(cur_row,NEWBTS    );
   table[cur_row][NEWLAC    ]=cells_ex(cur_row,NEWLAC    );
   table[cur_row][NEWCI     ]=cells_ex(cur_row,NEWCI     );
   table[cur_row][NEWET     ]=cells_ex(cur_row,NEWET     );
   table[cur_row][NEWNSEI   ]=cells_ex(cur_row,NEWNSEI   );   
   table[cur_row][NEWDAP   ]=cells_ex(cur_row,NEWDAP   );
   cur_row=cur_row+1;
   total_row=total_row+1;
   bsc_s=cells_ex(cur_row,1);
   //stop();
   endwhile
   table[cur_row][OLDBSC    ]="";
	 table[cur_row][OLDBCF    ]="";
	 table[cur_row][OLDSEG    ]="";
	 table[cur_row][OLDBTS    ]="";
	 table[cur_row][OLDLAC    ]="";
	 table[cur_row][OLDCI     ]="";
	 table[cur_row][OLDET     ]="";
	 table[cur_row][NEWBSC    ]="";
	 table[cur_row][NEWBCF    ]="";
	 table[cur_row][NEWSEG    ]="";
	 table[cur_row][NEWBTS    ]="";
	 table[cur_row][NEWLAC    ]="";
	 table[cur_row][NEWCI     ]="";
	 table[cur_row][NEWET     ]="";	 
	 table[cur_row][NEWNSEI   ]="";
	 table[cur_row][NEWDAP    ]="";
	 
endfunction

integer function getnewci_row(string p_ci,string p_lac)      
integer i,j;                                                 
j=arraysize(celllist);                                       
    for (i=1;i<=j;i++)                                       
      if (p_ci==celllist[i])                                 
          if  (p_lac==laclist[i])                            
             return i+1;                                     
          endif                                              
      endif                                                  
    endfor                                                   
    return 0;                                                
endfunction                                                  

string function mytrim(string x)
x=strtrim(x,1);
x=strtrim(x,2);
return x;
endfunction 

string function myif(integer x,string y,string z)
if (x<>0)
return y;
else 
return z;
endif
endfunction 

integer function mystoi(string x)
return strtoint(x,DEC);
endfunction






function CreateADJ()
string pmrg,lmrg,qmrg,sl,sync,drt,rac,agena,POPT,TRHO,aucl,grxp,acl,fmt;
integer ncc,bcc,freq,pmax1,pmax2;
line="";
 while ( NOT strstr(line,"COMMAND EXECUTED"))
             getline(line);
	if (strstr(line,"LOCATION AREA CODE......")) 
			 		///stop();
			 			lac=mymid(line,49,6);
			 			getline(line);
			 			ci=mymid(line,49,6);
			 			new_lac_row=getnewci_row(ci,lac);
			 			if (new_lac_row >0)
				 			lac=cells(new_lac_row,NEWLAC);
				 			//to add different ci here
			 			endif
			 		//	getlineback(line);
			 		endif
			 		
			 		if (strstr(line,"...(NCC).."))
			 		  ncc=strtoint(mymid(line,49,6),DEC);
			 		  
			 		endif
			 		
			 		if (strstr(line,"...(BCC).."))
			 		   bcc=strtoint(mymid(line,49,6),DEC);
			 		endif
			 		
			 		if (strstr(line,"...(FREQ).."))
			 		   freq=strtoint(mymid(line,49,6),DEC);
			 		endif
			 		
			 		
			 		if (strstr(line,"(PMRG)"))   
               pmrg=mymid(line,49,5); 
          endif     
          
          if (strstr(line,"(LMRG)"))   
             lmrg=mymid(line,49,5);   
          endif      
          
          if (strstr(line,"(QMRG)"))   
             qmrg=mymid(line,49,5);   
          endif      
			 		
			 		if (strstr(line,"(SL)"))   
             sl=mymid(line,50,4);
             getline(line);
             aucl=mymid(line,50,4);  
          endif
                                

			 		if (strstr(line,"(PMAX1)"))
			 		  pmax1=strtoint(mymid(line,49,5),DEC);
			 		  getline(line);
			 		  pmax2=strtoint(mymid(line,49,5),DEC);
			 		endif
			 		
			 		if (strstr(line,"(SYNC)"))
			 		  sync=mymid(line,52,1);
			 		endif
			 		
			 		if (strstr(line,"..(TRHO)."))
			 		  TRHO=myif(mymid(line,50,4)=="N","N","-"+mymid(line,50,4));
			 		  
			 		endif
			 		
			 		if (strstr(line,"..(ACL)."))
			 		  acl=myif(mymid(line,50,4)=="N","N",mymid(line,52,5));
			 		  
			 		endif
			 		
			 	if (strstr(line,".(FMT)."))
			 		  fmt=myif(mymid(line,49,5)=="0",",FMT="+mymid(line,49,5),",FMT="+mymid(line,49,5));
			 		  
			 		endif
			 		
			 		
			 		
			 		if (strstr(line,"...(POPT)"))
			 		  POPT=myif(mymid(line,50,4)=="N","N","-"+mymid(line,50,4));
			 		endif
			 		
			 		if (strstr(line,"(DRT)")) 
			 		  drt=mymid(line,50,4);   
			 		endif                      
			 	
			 	if (strstr(line,"(AGENA)"))
			 		  agena=mymid(line,52,1);
			 		  getline(line);
			 		  grxp=mymid(line,50,4);
			 		endif	
			 		
			 		//ZEAC:BTS=470::LAC=29795,CI=58002:NCC=3,BCC=1,FREQ=728:SYNC=N,SL=-95,PMRG=5,QMRG=2,LMRG=3,DRT=-85,AGENA=Y,RAC=1;			 		 		
			 		//if ((strstr(line,"(DTM)"))OR (strstr(line,"(DHOPWR)"))) //如果需要在这里修改相邻小区的最后一个参数的标记
			 			if ((strstr(line,".....(RAC).")))		
			 			
			 			rac=mymid(line,49,5);
			 			
			 			 			 					 			
			 			fileprint(adj_file,"ZEAC:SEG=%s::LAC=%s,CI=%s:NCC=%d,BCC=%d,FREQ=%d:SYNC=%s,TRHO=%s,ACL=%s%s,POPT=%s,SL=-%s,AUCL=-%s,PMAX1=%d,PMAX2=%d,PMRG=%s,QMRG=%s,LMRG=%s,DRT=-%s,AGENA=%s,GRXP=-%s,RAC=%s;\n",cells_ex(cur_row,NEWSEG),lac,ci,ncc,bcc,freq,sync,TRHO,acl,fmt,POPT,sl,aucl,pmax1,pmax2,pmrg,qmrg,lmrg,drt,agena,grxp,rac);
			 		//	lac=""; ci="";			 			
			 		endif
			 		
			  endwhile	

endfunction


function Create_3GADJ()
integer MCC,RNC,CI,LAC,SAC,FREQ,SCC;
string  MNC;
line="";
 while ( NOT strstr(line,"COMMAND EXECUTED"))
             getline(line);
	//if (strstr(line,"HAS WCDMA RAN ADJACENT CELL")) 
		//	 			SEG=strtoint(mymid(line,5,4),DEC);
			// 		endif
			 		
			 		if (strstr(line,"MOBILE COUNTRY CODE....................(MCC)...."))
			 		  MCC=strtoint(mymid(line,50,3),DEC);  
			 		endif
			 		
			 		if (strstr(line,"MOBILE NETWORK CODE....................(MNC)...."))   
             MNC=mymid(line,50,3);   
          endif     
          
          if (strstr(line,"RADIO NETWORK CONTROLLER...............(RNC)...."))   
             RNC=strtoint(mymid(line,50,6),DEC);   
          endif      
          
          if (strstr(line,"CELL IDENTIFICATION....................(CI)....."))   
             CI=strtoint(mymid(line,50,6),DEC);   
          endif      
			 		
			 		if (strstr(line,"LOCATION AREA CODE.....................(LAC)...."))   
             LAC=strtoint(mymid(line,50,6),DEC);   
          endif
                                

			 		if (strstr(line,"SERVICE AREA CODE......................(SAC)...."))
			 		  SAC=strtoint(mymid(line,50,6),DEC);
			 		  
			 		endif
			 		
			 		if (strstr(line,"WCDMA DOWNLINK CARRIER FREQUENCY.......(FREQ)..."))
			 		  FREQ=strtoint(mymid(line,50,6),DEC);
			 		endif
			 		
			 		if (strstr(line,"SCRAMBLING CODE........................(SCC)....")) 
			 		  SCC=strtoint(mymid(line,50,6),DEC);  
		              
		              
		   
		                 
			// 	stop();
			fileprint(adj3G_file, "ZEAE:SEG=%s::MCC=%d,MNC=%s,RNC=%d,CI=%d:LAC=%d,SAC=%d,SCC=%d,FREQ=%d;\n",cells_ex(cur_row,NEWSEG),MCC,MNC,RNC,CI,LAC,SAC,SCC,FREQ);
//fileprint(adj_file,"ZEAC:SEG=%s::LAC=%s,CI=%s:NCC=%d,BCC=%d,FREQ=%d:SYNC=%s,SL=-%s,PMAX1=%d,PMAX2=%d,PMRG=%s,QMRG=%s,LMRG=%s,DRT=-%s,AGENA=%s,RAC=%s;\n",cells_ex(cur_row,NEWSEG),lac,ci,ncc,bcc,freq,sync,sl,pmax1,pmax2,pmrg,qmrg,lmrg,drt,agena,rac);		 			
			 		endif
			 		
			  endwhile	

endfunction


function Create_TDADJ()
string rnc,ci,lac,sac,freq,cpa,div,sc;
line="";
 while ( NOT strstr(line,"COMMAND EXECUTED"))
             getline(line);
	if (strstr(line,"RADIO NETWORK CONTROLLER...............(RNC)....")) 
			 			rnc=mymid(line,50,5);
			 		endif
			 		
 if (strstr(line,"CELL IDENTIFICATION....................(CI)...."))
			 		  ci=mymid(line,50,5);
			 		  getline(line);
			 		  lac=mymid(line,50,5);
			 		  getline(line);
			 		  sac=mymid(line,50,5);
			 		  getline(line);
			 		  freq=mymid(line,50,5);
			 		  getline(line);
			 		  cpa=mymid(line,50,5);
			 			getline(line);
			 		  div=mymid(line,50,5); 
			 			getline(line);
			 		  sc=mymid(line,50,2);  
			 			 
			 		fileprint(adjTD_file,"ZEAF:SEG=%s::MCC=460,MNC=00,RNC=%s,CI=%s:LAC=%s,SAC=%s,FREQ=%s,CPA=%s,DIV=%s,SC=%s;\n",cells_ex(cur_row,NEWSEG),rnc,ci,lac,sac,freq,cpa,div,sc);	 				 					 			
			 	//fileprint(adj_file,"ZEAF:SEG=%s::LAC=%s,CI=%s:NCC=%d,BCC=%d,FREQ=%d:SYNC=%s,SL=-%s,PMAX1=%d,PMAX2=%d,PMRG=%s,QMRG=%s,LMRG=%s,DRT=-%s,AGENA=%s,RAC=%s;\n",cells_ex(cur_row,NEWSEG),lac,ci,ncc,bcc,freq,sync,sl,pmax1,pmax2,pmrg,qmrg,lmrg,drt,agena,rac);
		 			
			 		endif
			 		
			  endwhile	

endfunction

/////////Event Define  ///////////////////////////////////////////
defevent RESPONSE
    id:         "DX_ERROR_RESPONSE";
    initstat:   ACTIVE;
    respat:     "'DX ERROR:'   ";
    resline:    ERRORLINE;
    call:       DX_ERROR_RESPONSE(ERRORLINE);
enddefevent

defevent RESPONSE
    id:         "ERROR_RESPONSE_1";
    initstat:   ACTIVE;
    respat:     "'UNKNOWN COMMAND GROUP'            OR\
                 'UNKNOWN COMMAND CLASS'            OR\
                 'COMMAND EXECUTION FAILED'         OR\
                 'REQUESTED FUNCTION NOT'           OR\
                 'EXCESSIVE INPUT DELAY'            OR\
                 'REMOVING NOT POSSIBLE'            OR\
                 'DISK NOT UPDATED'                 OR\
                 'MODIFICATION NOT SUCCEEDED'       OR\
                 'COMMAND EXECUTION ABORTED'";
    resline:    ERRORLINE;
    call:       ERROR_RESPONSE(ERRORLINE);
enddefevent

/*define event when there is line containing one of strings in the respat
  section */

defevent RESPONSE
    id:         "ERROR_RESPONSE_2";
    initstat:   ACTIVE;
    respat:     "'MML PROGRAM LOAD ERROR'           OR\
                 'COMMAND NOT AUTHORIZED'           OR\
                 'COMMAND CHECK ERROR'              OR\
                 'DXPFIL WRITE ERROR'               OR\
                 'DXPFIL READ ERROR'                OR\
                 'NO RESPONSE FROM VIDAST'          OR\
                 'SYNTAX ERROR'";
    resline:    ERRORLINE;
    call:       ERROR_RESPONSE(ERRORLINE);
enddefevent

/*define event when there is line containing "DEBUGGER COMMAND..."
  in the response.*/

defevent RESPONSE
    id:         "ERROR_RESPONSE_3";
    initstat:   ACTIVE;
    respat:     "'DEBUGGER COMMAND EXECUTION ERROR'";
    resline:    ERRORLINE;
    call:       ERROR_RESPONSE(ERRORLINE);
enddefevent

defevent RESPONSE
        id:          "ERROR_RESPONSE_4";
        initstat:    ACTIVE;
        respat:      "'COMMAND CURRENTLY NOT ALLOWED' OR\
        					 '*** USER AUTHORIZATION FAILURE'";
        resline:     ERRORLINE;
        call:        ERROR_RESPONSE(ERRORLINE);
enddefevent

/* ERROR_RESPONSE_5 is deactivated before ZAHO -command and activated
   again after command is executed */

defevent RESPONSE
    id:         "ERROR_RESPONSE_5";
    initstat:   ACTIVE;
    respat:     "'SYSTEM ERROR'";
    resline:    ERRORLINE;
    call:       ERROR_RESPONSE(ERRORLINE);
enddefevent

/* ERROR_RESPONSE_6 is deactivated before ZGCF -command and activated
   again after command is executed */

defevent RESPONSE
    id:         "ERROR_RESPONSE_6";
    initstat:   ACTIVE;
    respat:     "'SEMANTIC ERROR'";
    resline:    ERRORLINE;
    call:       ERROR_RESPONSE(ERRORLINE);
enddefevent

handler ERROR_RESPONSE(string eline)
    // Return with error status
    messagebox(strprint("Erroneus response: %s, please check before"
                        " continuing.",eline),MB_ERROR);
    erreturn;
endhandler


handler DX_ERROR_RESPONSE(string eline)
        strscan(eline,"%*[^:]:%s",eline);
      switch(eline)
        case "11266"://BCC+NCC+BCCH TRX FREQUENCY NOT UNIQUE IN ADJACENCY DEFINITIONS
			break;
		  case "10201"://OBJECT ALREADY EXISTS IN DATABASE 
		  	break;	
		  case "10548":// MAXIMUM NUMBER OF ADJACENT 
		   break;  
        default:
        /* show messagebox and return to line which caused error */
        messagebox(strprint("Erroneus response: %s, please check before"
                            " continuing.",eline),MB_ERROR);
        erreturn;    
      endswitch
		return;
endhandler


  