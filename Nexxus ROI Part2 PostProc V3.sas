*********************************************************************************************
					NEXXUS Process Part-2
	This is the ROI analysis part. It starts from the matched pair created in Part-1 process.
	This code will run ANCOVA for product x, competitor products and market Rx and 
	prepare output for all designed metrics. THe out put tables will be written in excel file 
	with tabs:
		Attrition, Descriptive_Table, Market Penetration, ANCOVA_Test, PreMatch_Test_Ctrl
		and PostMatch_Test_Ctrl, Monthly_Rx (For pre-post Trend)

 	To properly run this code you need assign the correct lib names and data locations as shown below:
		libname sasdatax ... 
		%let Dat_dir = ...
		%let Out_dir = ...

	Final excel file will be saved in Out_dir with file name: OutPut_Campaign_&Prog_num..xlsx

*********************************************************************************************;
libname sasdatax '\\plyvnas01\statservices2\customstudies\Promotion Management\2014\NOVOLOG\03 Data';
%let Dat_dir = C:\work\working materials\SAS to R;
%let Out_dir = \\plyvnas01\statservices2\customstudies\Promotion Management\2014\NOVOLOG\05 Reports;
*********************************************************************************************;
proc import out=work.matched_data
datafile="\\plyvnas01\StatServices2\CustomStudies\Promotion Management\2014\NOVOLOG\04 Codes\Nexxus\R_Version\1016\hcp_matched_all.csv"
dbms=csv replace;
run;
 %let Raw_Xpo_data = for_sas_test_0929; /* Xponent Raw Rx Data*/
				%let  Campaign_PRD = VICTOZA ;/* Campaign Product Name */ 
				 %let Pre_wks	  = 26   ;   /* Same as defined in part-1 */
			%let 	 Target_RX	  = NRX   ;  /* Match and analysis Rx. Same as used in Part-1 */
			%let 	 Total_wks    = 128    ; /* Total available data weeks. Same as used in Part-1 */
			%let 	 Campaign_dat = CAMPAIGN 139; /* Campaign Contact HCP list excel file */
			%let 	 Specfile	  = SPEC_Grp     ;  /* Excel file for customized specialty groups. Default = All in 1 */
			%let 	 ZIP2Region	  = ZIP_region    ;     /* Excel file for Zip to customized region. Defaulat=Census regions */

%let 	 Specfile	  = .     ;  /* Excel file for customized specialty groups. Default = All in 1 */
			%let 	 ZIP2Region	  = .    ;     /* Excel file for Zip to customized region. Defaulat=Census regions */

	
%macro Nexxus_ROI_main(
				 Raw_Xpo_data = for_sas_test_0929, /* Xponent Raw Rx Data*/
				 Campaign_PRD = VICTOZA, /* Campaign Product Name */ 
				 Pre_wks	  = 26,      /* Same as defined in part-1 */
				 Target_RX	  = NRX,     /* Match and analysis Rx. Same as used in Part-1 */
				 Total_wks    = 128,     /* Total available data weeks. Same as used in Part-1 */
				 Campaign_dat = CAMPAIGN 139, /* Campaign Contact HCP list excel file */
				 Specfile	  = . ,        /* Excel file for customized specialty groups. Default = All in 1 */
				 ZIP2Region	  = .        /* Excel file for Zip to customized region. Defaulat=Census regions */);
				
	%global prog_num;
	%let Campaign_Contact =%sysfunc(tranwrd(%sysfunc(tranwrd(&Campaign_dat,%str( ),%str(_))),%str(.),%str(_)));

	* Get campaign program ID from data ;
	proc sql; select campaign_ID into :prog_num from  sasdatax.&Campaign_Contact._ID; quit;

	*** Readin Specialty grouping definition, if provided ***;
	%if ".&Specfile" ne "." %then %do; 
		proc import datafile= "&dat_dir./&Specfile..csv" dbms=csv out=MY_SPEC replace;
			
		run;

		DATA Specok;  set MY_SPEC;
			Retain fmtname '$SpecOk';
			start = spec_cd;
			label = "OK";
		PROC FORMAT CNTLIN=Specok; RUN;
		DATA MYSPEC;  set MY_SPEC;
			Retain fmtname '$SPECGRP';
			start = spec_CD;
			label = class_desc;
		PROC FORMAT CNTLIN=MYSPEC; RUN; 
	%end;
	*** Readin Region definition,  if provided ***;
	%if ".&ZIP2Region" ne "." %then %do; 
		proc import datafile= "&dat_dir./&ZIP2Region..csv" dbms=csv out=MY_Rgn replace;
			getnames = yes;
		run;

		data MY_Rgn;
		  set MY_Rgn;
/*		  zipcode1 = put(zipcode, z5.);*/
zipcode1 = strip(put(zipcode, $7.));
drop zipcode;
		  rename zipcode1 = zipcode;
		run;

		DATA zipok;  set MY_Rgn;
			Retain fmtname '$ZIPOk';
			start = ZIPCODE;
			label = "OK";
		PROC FORMAT CNTLIN=zipok; RUN;
		DATA MYRGN;  set MY_Rgn;
			Retain fmtname '$MYRGN';
			start = ZIPCODE;
			label = Region;
		PROC FORMAT CNTLIN=MYRGN; RUN;
	%end;

	* Get required parameters from input data;
	proc sql; select min(Engagement_date) into :CStart from sasdatax.Campaign_HCP_&Prog_num.; quit;
	data _null_; Set sasdatax.&Raw_Xpo_data (obs=1); 
		start_wk = &total_wks. + week(&CStart)+(year(&CStart)-year(Datadate))*52 -week(Datadate) - 26; 
		End_wk   = &total_wks. + week(&CStart)+(year(&CStart)-year(Datadate))*52 -week(Datadate) - 1; 
		call symput("IMSDate2", Datadate);
		call symput("from",trim(left(put(start_wk,3.))));
		call symput("to",trim(left(put(End_wk,3.))));
		*put Datadate date10.;
		
	run;

%put &from;
%put &to;
	data TC_&Prog_num._all ; 
		set matched_data (keep=IMSDR7 Engagement_Date cohort Engage_wk TRX_oth_1-TRX_oth_&Total_wks.
		                               TRX_wk_1-TRX_wk_&Total_wks. NRX_oth_1-NRX_oth_&Total_wks. NRX_wk_1-NRX_wk_&Total_wks.);
		array othT TRX_oth_1-TRX_oth_&Total_wks.;
		array othN NRX_oth_1-NRX_oth_&Total_wks.;
		array nlT TRX_wk_1-TRX_wk_&Total_wks.;
		array nlN NRX_wk_1-NRX_wk_&Total_wks.;

		retain pre_period_nlT post_period_nlT pre_period_othT post_period_othT
		       pre_period_nlN post_period_nlN pre_period_othN post_period_othN;
		pre_period_nlT=0; post_period_nlT=0; pre_period_othT =0; post_period_othT=0;
		pre_period_nlN=0; post_period_nlN=0; pre_period_othN =0; post_period_othN=0;
		n_pre=0; n_post = 0;
		if cohort = 1 then group="TEST"; else group="CONTROL";

		do i = 1 to &Total_wks.;
			gap = i-Engage_wk-1;
		  if gap > (-1-&pre_wks.) and gap < 0 then do;
			  n_pre = n_pre + 1;
				pre_period_nlT = sum(pre_period_nlT , nlT(i)); /* treat missing values as zeros */
				pre_period_othT = sum(pre_period_othT , othT(i)); /* treat missing values as zeros */
				Pre_Market_T = pre_period_nlT + pre_period_othT ;

				pre_period_nlN = sum(pre_period_nlN , nlN(i)); /* treat missing values as zeros */
				pre_period_othN = sum(pre_period_othN , othN(i)); /* treat missing values as zeros */
				Pre_Market_N = pre_period_nlN+pre_period_othN ;
			end;
			if gap > 0  then do;
				n_post = n_post + 1;
				post_period_nlT = sum(post_period_nlT , nlT(i)); /* this increase the test HCPs from 1,034 to 2,541 */
				post_period_othT = sum(post_period_othT , othT(i)); /* this increase the test HCPs from 1,034 to 2,541 */
				Post_Market_T = post_period_nlT+post_period_othT ;

				post_period_nlN = sum(post_period_nlN , nlN(i)); /* this increase the test HCPs from 1,034 to 2,541 */
				post_period_othN = sum(post_period_othN , othN(i)); /* this increase the test HCPs from 1,034 to 2,541 */
				Post_Market_N = post_period_nlN+post_period_othN ;
			end;
		end;
		* Standarzizing pre and post period data to reflect same months of data;
		pre_period_nlT  = pre_period_nlT *&pre_wks./n_pre;
		pre_period_othT = pre_period_othT*&pre_wks./n_pre;
		post_period_nlT = post_period_nlT*&pre_wks./n_post;
		post_period_othT= post_period_othT*&pre_wks./n_post; 
		pre_period_nlN  = pre_period_nlN *&pre_wks./n_pre;
		pre_period_othN = pre_period_othN*&pre_wks./n_pre;
		post_period_nlN = post_period_nlN*&pre_wks./n_post;
		post_period_othN= post_period_othN*&pre_wks./n_post; 

		Pre_Market_T     = Pre_Market_T*&pre_wks./n_pre; 
		Post_Market_T    = Post_Market_T*&pre_wks./n_post; 
		Pre_Market_N     = Pre_Market_N*&pre_wks./n_pre; 
		Post_Market_N    = Post_Market_N*&pre_wks./n_post; 

		keep IMSDR7 pre_period_: post_period_: group TRX_oth_1-TRX_oth_&Total_wks. TRX_wk_1-TRX_wk_&Total_wks. 
		     NRX_oth_1-NRX_oth_&Total_wks. NRX_wk_1-NRX_wk_&Total_wks. Engagement_Date 
			 gap cohort Pre_Market_: Post_Market_:; 
	run;
	* Get HCP counts for matched Test and controls after check min post weeks ;
	Proc sql; 
    Select count(distinct IMSDR7) into :Line6 from TC_&Prog_num._all(where=(Cohort = 1)); 
    select count(distinct imsdr7) into :Line7 from TC_&Prog_num._all(where=(Cohort = 0)); 
  quit;

	* Get the adjusted Post RX counts using Proc GLM or MIXED - the ANCOVA analysis *;
	proc mixed data = TC_&Prog_num._all;
	  class group;
		model post_period_nlN = group pre_period_nlN/ SOLUTION;
		lsmeans group / DIFFS ;
	 ods output LSMeans = LSM_X DIFFs = Diff_x;
	run;
	proc transpose data = lsm_x out=lsm_xt;
		id group;
		var estimate;
	run;
	data _null_; set Diff_x; call symput("px", probt); run;
	proc mixed data = TC_&Prog_num._all;
	  class group;
		model Post_Market_n  = group Pre_Market_n / SOLUTION;
		lsmeans group/DIFFS   ;
	  ods output LSMeans = LSM_M DIFFs = Diff_M;
	run;
	proc transpose data = lsm_m out=lsm_mt;
		id group;
		var estimate;
	run;
	data _null_; set Diff_m; call symput("pm", probt); run;
	proc mixed data = TC_&Prog_num._all;
	  class group;
		model post_period_othN = group pre_period_othN/ SOLUTION;
		lsmeans group/ DIFFS  ;
	 ods output LSMeans = LSM_C DIFFs = Diff_C;
	run;
	proc transpose data = lsm_c out=lsm_ct;
		id group;
		var estimate;
	run;
	data _null_; set Diff_c; call symput("pc", probt); run;
	data ancova_Test; retain Product Matched_HCPs Test cont Gains Index P_Value;
		set lsm_Xt(in=a) lsm_mt(in=b) lsm_Ct(in=c);
		Gains = Test-Control;
		if a then do; Product = "Product X            "; Matched_HCPs=&line6.; Index = Test/cont; P_Value = &px. ; end;
		if b then do; Product = "Market (All Products)"; Matched_HCPs=&line6.; Index = Test/cont; P_Value = &pm. ; end;
		if c then do; Product = "Product Other than X "; Matched_HCPs=&line6.; Index = Test/cont; P_Value = &pc. ; end;
		drop _name_;
	run;
	/* END of ANCOVA Analyze Output data*/
	  
	/* START: create points for graph */
	data graph_&Prog_num. ; length SPECIALTY $20.;
/*		set sasdatax.HCP_MATCHED_ALL_&Prog_num.;*/
	set matched_data;
/*		zipcode_num= input(zipcode, 8.0);*/ /*character to numeric*/
/*		drop zipcode;*/
/*		rename zipcode_num=zipcode;*/

/*	zipcode_char=put(zipcode, $8.);*/
/*	drop zipcode;*/
/*	rename zipcode_char=zipcode;*/
	zipcode1=strip(zipcode);
	drop zipcode;
	rename zipcode1=zipcode;

		array othT TRX_oth_1-TRX_oth_&Total_wks.;
		array othN NRX_oth_1-NRX_oth_&Total_wks.;
		array nlT TRX_wk_1-TRX_wk_&Total_wks.;
		array nlN NRX_wk_1-NRX_wk_&Total_wks.;

		if cohort = 1 then Group="TEST"; else Group="CONTROL";
/*		%if ".&ZIP2Region" ne "." %then %do; */
			if put(Spec,$specok.) = "OK" then SPECIALTY = put(Spec, $SPECGRP.); else SPECIALTY = "ALL OTHER";
/*		%end;*/
/*		%else %do; */
/*      SPECIALTY = "ALL SPECIALTY";*/
/*		%end;*/
/*		%if ".&ZIP2Region" ne "." %then %do; */
			flag_rg = put(ZIPCODE, $ZIPOk.);
flag_spec = put(spec, $specok.);
			Region  = put(ZIPCODE, $MYRGN.);
			specialty=put(spec, $specgrp.);
			if put(ZIPCODE,$ZIPOk.) = "OK" then Region  = put(ZIPCODE, $MYRGN.); else Region = "OTHER";
/*		%end;*/
		do i = 1 to &Total_wks. ;
			gap = i-Engage_wk-1;
	    if gap > (-1-&pre_wks.) and gap < (1+&pre_wks.)  then do;
				if gap > 0 then post_n = Post_n + 1;
			  rel_wk=gap;
				if -&pre_wks    <=gap < -&pre_wks+4  then Rel_month = -6;
				if -&pre_wks+4  <=gap < -&pre_wks+8  then Rel_month = -5;
				if -&pre_wks+8  <=gap < -&pre_wks+13 then Rel_month = -4;
				if -&pre_wks+13 <=gap < -&pre_wks+17 then Rel_month = -3;
				if -&pre_wks+17 <=gap < -&pre_wks+21 then Rel_month = -2;
				if -&pre_wks+21 <=gap < -&pre_wks+26 then Rel_month = -1;
				*if gap>0 then rel_month= ceil((gap+0)/4);
				if 0  <=gap < 4  then Rel_month = 1;
				if 4  <=gap < 8  then Rel_month = 2;
				if 8  <=gap < 13 then Rel_month = 3;
				if 13 <=gap < 17 then Rel_month = 4;
				if 17 <=gap < 21 then Rel_month = 5;
				if 21 <=gap < 26 then Rel_month = 6;
				************************************************;
				if gap>=-&pre_wks. and gap <=-14 then Period = -3; 
				else if gap>=-13 and gap <=-5 then Period = -2; 
				else if gap>=-4 and gap <=-1 then Period = -1; 
				else Period = rel_month;
				***********************************************;
				rel_trx=nlT(i);      rel_Nrx=nlN(i);
				rel_trx_oth=othT(i); rel_Nrx_oth=othN(i);
				MKT_TRX = rel_trx+rel_trx_oth; MKT_NRX = rel_Nrx+rel_Nrx_oth;
				output;
			end;
		 end;
		keep flag_rg flag_spec IMSDR7 group rel_wk rel_month rel_trx rel_trx_oth rel_Nrx rel_Nrx_oth MKT_TRX MKT_NRX Period specialty Region; 
	run;
	***********  Get Monthly summary ****************;
	proc means data =graph_&Prog_num. noprint nway ;
		class Group rel_month;
		var rel_trx  rel_trx_Oth MKT_TRX rel_Nrx  rel_Nrx_Oth MKT_NRX;
		output out= graph_data(drop= _type_ rename=(_freq_=N)) 
		/*mean=Prod_TRx_Mean Oth_TRX_Mean MKT_TRX_Mean Prod_NRx_Mean Oth_NRX_Mean MKT_NRX_Mean*/
		sum=ProdX_TRx Oth_TRx MKT_TRX ProdX_NRx Oth_NRx MKT_NRX;
	run;
	proc sql;
		create table HCP_CNT as select group,rel_month, count(distinct IMSDR7) as HCPs from graph_&Prog_num. 
			   group by Group,rel_month order by group,rel_month; quit;
			   proc print data=HCP_CNT; run;
	proc sort data=graph_data; by group rel_month; run;
	data graph_data;
		retain Group Rel_month HCPs ProdX_NRx Oth_NRx MKT_NRX ProdX_TRx Oth_TRx MKT_TRX ;
		merge graph_data HCP_CNT; 
		by group rel_month;
		keep Group Rel_month HCPs ProdX_NRx Oth_NRx MKT_NRX ProdX_TRx Oth_TRx MKT_TRX ;
	run;
		 
	***********  Get Post Match test/control comparison ****************;
	proc means data =graph_&prog_num. noprint nway missing; where Period < 0;
		class specialty Region group IMSdr7 Period;
		var rel_trx rel_trx_Oth MKT_TRX rel_Nrx rel_Nrx_Oth MKT_NRX;
		output out= Monthly_data sum=;
	run;
	proc means data = Monthly_data missing noprint ;
		class group specialty Region Period;
		var rel_trx rel_trx_Oth MKT_TRX rel_Nrx rel_Nrx_Oth MKT_NRX;
		output out= Monthly_mean (drop=_type_) mean= std=Nov_STD Oth_STD MKT_STD Nov_STD2 Oth_STD2 MKT_STD2;
	run;
	data Monthly_mean; set Monthly_mean;
		where Period ne . and group ne " " and specialty ne " ";
		*if specialty = " " then specialty = "ALL";
		if Region = " " then Region = "ALL ";
	proc sort data = Monthly_mean; 
		by specialty Region Period _FREQ_; 
	run;
	proc transpose data = Monthly_mean out=Pre_Compare; by specialty Region Period ;
		id group;
		var _FREQ_ rel_trx rel_trx_Oth MKT_TRX rel_Nrx rel_Nrx_Oth MKT_NRX Nov_STD Oth_STD MKT_STD Nov_STD2 Oth_STD2 MKT_STD2; 
	run;
	proc sort data = Pre_Compare out=nov_RX(keep=specialty Region Period CONT TEST);
		where UPCASE(_name_) = "REL_NRX"; by specialty Region Period; run;
	proc sort data = Pre_Compare out=nov_std(keep=specialty Region Period CONT TEST  
		rename=(CONT=C_STD TEST=T_STD)); 
		where UPCASE(_name_) = "NOV_STD2"; by specialty Region Period; run;
	proc sort data = Pre_Compare out=freq(keep=specialty Region Period CONT TEST
		rename=(CONT=N_CTRL TEST=N_TEST)); 
		where UPCASE(_name_) = "_FREQ_"; by specialty Region Period; run;
	data Post_match_Compare;
		retain SPECIALTY Region Period N_TEST N_CTRL CONT TEST C_STD T_STD P_Value;
		merge nov_RX nov_std freq;
		by specialty Region Period;
		if T_STD>0 then 
		P_Value = 2*(1-PROBNORM(ABS(TEST-CONT)/SQRT(T_STD**2/N_TEST+ C_STD**2/N_CTRL)));
		else P_Value = .;
	run;
	*proc print data = Post_match_Compare; run;

	***********  Final Output **************;
	proc export data = sasdatax.lines1_&prog_num. outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace; 
		sheet="Attrition" ; run;
	proc export data = ancova_Test outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace ; 
		sheet="Ancova_Test" ;  run;
	proc export data =sasdatax.descriptives outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace  ; 
		sheet="Descriptive Table"; run;
	proc export data =sasdatax.PENETRATION   outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace  ; 
		sheet="Market Penetration"; run;
	proc export data =sasdatax.Pre_Match_Compare outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace  ; 
		sheet="PreMatch Test_Ctrl"; run;
	proc export data =Post_match_Compare outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace  ; 
		sheet="PostMatch Test_Ctrl"; run;
	proc export data = graph_data outfile = "&Out_dir.\OutPut_Campaign_&Prog_num..xlsx" dbms=excel replace  ; 
		sheet="Monthy_Rx" ; run;
%mend Nexxus_ROI_main;

%Nexxus_ROI_main(Raw_Xpo_data = ims_rx_data_sample,
				 Campaign_PRD = VICTOZA,
				 Pre_wks = 26,
				 Target_RX = NRX,
				 Total_wks    = 128,
				 Campaign_dat = CAMPAIGN 139
				);

**************** END of PROCESS ****************;
 
