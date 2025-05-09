library(xlsx)
library(emayili)
library(mailR)
pricingdate<-as.Date("2025-05-08") ##估值日期
sender<- "fdi_trade_service@xyzq.com.cn"   ##"fdi_trade_service@xyzq.com.cn"
emailbody<-"请以此为准，谢谢！
您好，估值文件如附，请查收，如对估值报告有疑问，请联系：fdi_trade_service@xyzq.com.cn，谢谢！
  
免责声明：

1、兴业证券股份有限公司（简称“兴业证券”）依据商业合理原则，基于我方内部模型计算得出相关衍生品合约于相应日期的估值报告（统称为“估值报告”），该估值报告计算结果仅供贵方参考，并非兴业证券对贵方所持相关衍生品合约权益价值的保证，亦不代表兴业证券对贵方是否应当追加保证金（包括追加预付金等）以及追加保证金金额作出任何意思表示。若贵方在相关衍生品合约下需追加保证金，兴业证券将另行通知贵方。

2、估值报告基于兴业证券内部模型计算得出，该内部模型的调整、优化、参数使用等由兴业证券依据商业合理原则自行决定。

3、估值报告并不构成兴业证券对贵方有关任何事项（包括但不限于相关衍生品合约的提前终止价格、价值以及计算基准、计算公式）的要约或承诺。若贵方拟提前终止相关衍生品合约，应当另行向兴业证券询价。

4、请贵方审慎核对该估值报告。兴业证券不对贵方或任何第三方使用估值报告产生的结果承担任何责任。如对本估值报告有疑问，请及时回复该邮箱进行咨询、反馈。

5、若估值报告因政策法规变化、自然灾害、系统故障、行情传输错误、行情不稳定等不可抗力出现延迟、中断、错误等，兴业证券不承担任何责任。

6、估值报告不应作为投资相关衍生品合约的资产管理产品（以下简称“资管产品”）进行任何份额或权益变更的依据。若资管产品基于估值报告载明的信息进行开放申购、赎回等份额或权益变更操作，兴业证券不对该等操作的合理性、公允性和法律适当性承担任何责任。若贵方与资管产品投资者就估值报告载明的估值产生任何争议，由贵方自行解决，兴业证券不因此对资管产品投资者承担任何责任。

7、贵方不得向除资管产品的外包服务机构、托管机构或投资顾问以外的第三方披露本估值报告。

8.估值报告的计算结果以交易不可发生提前终止为前提，若贵方确有临时提前终止交易需求，应另行向兴业证券询价，估值报告计算结果不作为提前终止报价的参考依据。"

modifytrade<-read.xlsx("optionlist.xlsx",8,encoding = "UTF-8")
n<-nrow(modifytrade)
customerlist<-unique(modifytrade$CustomerName)
m<-length(customerlist)
livingdays<-as.numeric(pricingdate-modifytrade$Startdate)
annualizedPV<-ifelse(modifytrade$ModifyType!=2,modifytrade$Premium - modifytrade$NotionalPrincipal*modifytrade$Premium_Percentage*livingdays/365,
                     modifytrade$NotionalPrincipal*(1+modifytrade$Fixedrate_Annual*livingdays/365))

for ( i in 1: m){
  customerindex<-which(modifytrade$CustomerName==customerlist[i])
  receiver<-unlist(strsplit(modifytrade$MailAddress[customerindex][1],split = ","))
  receiver<-c("fdi_derivs_trading@xyzq.com.cn",receiver)
  #receiver<-"lvjiawei@xyzq.com.cn"
  if (modifytrade$ModifyType[customerindex[1]] == 2){
    reportname<-paste("兴业证券",customerlist[i],"收益凭证估值报告",pricingdate,sep = "-")
    raw_report<-loadWorkbook(paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
    raw_sheets<-getSheets(raw_report)
    raw_sheet1<-raw_sheets[[1]]
    rawPVmatrix<-readColumns(raw_sheet1,startRow = 5,endRow = 5,startColumn = 1,endColumn = 16,header = FALSE)
    modifiedPV<-max(annualizedPV[customerindex],rawPVmatrix[,14])
    modifiedNAV<-round(modifiedPV/rawPVmatrix[,9],4)
    addDataFrame(data.frame(modifiedPV,modifiedNAV),raw_sheet1,row.names = F,col.names = F,startRow = 5,startColumn = 14)
    saveWorkbook(raw_report,paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
  } else
  {
    reportname<-paste("兴业证券",customerlist[i],"场外交易估值报告",pricingdate,sep = "-")
    raw_report<-loadWorkbook(paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
    raw_sheets<-getSheets(raw_report)
    raw_sheet1<-raw_sheets[[1]]
    raw_sheet2<-raw_sheets[[2]]
    modified_report<-loadWorkbook("D:/workingdirectory/valuationreport_options_modified.xlsx")
    modified_sheets<-getSheets(modified_report)
    modified_sheet1<-modified_sheets[[1]]
    modified_sheet2<-modified_sheets[[2]]
    cs1<-CellStyle(modified_report)+Font(modified_report,name = "Calibri")+
      Border(color = "Black",position = c("TOP","BOTTOM","LEFT","RIGHT"),
             pen = "BORDER_THIN")+Alignment(h="ALIGN_CENTER",wrapText = TRUE)
    cs2<-cs1+DataFormat("yyyy/m/d")
    cs3<-cs1+DataFormat("#,##0.00")
    cs4<-cs1+Font(modified_report,isBold = TRUE)
    cashsummary<-readColumns(raw_sheet1,startRow = 6,endRow = 6,startColumn = 1,endColumn = 5,header = FALSE)
    nrows<-raw_sheets[[2]]$getLastRowNum()+1 
    valuationmatrix<-readColumns(raw_sheet2,startRow = 3,endRow = length(customerindex)+2,startColumn = 1,endColumn = 19,header = FALSE)
    if (modifytrade$ModifyType[i]==1){
      valuationmatrix<-data.frame(valuationmatrix[,1:18],pmax(annualizedPV[customerindex],valuationmatrix[,18]),valuationmatrix[,19])
      addDataFrame(data.frame(customerlist[i]),modified_sheet1,row.names = F,col.names=F,startRow = 1,startColumn = 2,colStyle = list(`1`=cs4))
      addDataFrame(data.frame(pricingdate),modified_sheet1,row.names = F,col.names = F,startRow = 2,startColumn = 2,colStyle = list(`1`=cs2))
      addDataFrame(cashsummary,modified_sheet1,row.names = F,col.names = F,startRow = 6,startColumn = 1,colStyle = list(`1`=cs1,`2`=cs3,`3`=cs3,`4`=cs3,`5`=cs3))
      addDataFrame(valuationmatrix,modified_sheet2,row.names = F,col.names = F,startRow = 3,startColumn = 1)
      saveWorkbook(modified_report,paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
    } else
    {
      addDataFrame(data.frame(pmax(annualizedPV[customerindex],valuationmatrix[,18])),raw_sheet2,row.names = F,col.names = F,startRow = 3,startColumn = 18)
      saveWorkbook(raw_report,paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
    }
  }
  email<-envelope()
  email<-email%>% 
    from(sender)%>%
    to(receiver)
  email<-email%>%subject(reportname)%>%text(emailbody,charset = "utf-8")
  email<-email%>%attachment(paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""))
  
  smtp <- server(
    host = "mail.xyzq.com.cn",
    port = 25,
    username = sender,
    password = "BkgChK4BgKkXnKre",use_ssl = FALSE
  )
  smtp(email, verbose = TRUE)
  
  
  #send.mail(from = sender,to = receiver,
   # subject = reportname,body = emailbody,
   #smtp = list(host.name="mail.xyzq.com.cn",port=25,
         #      user.name=sender,
      #         passwd="BkgChK4BgKkXnKre",tsl=T),
  # authenticate = TRUE,send = FALSE,attach.files=
   # paste("D:/workingdirectory/ValuationFile/",reportname,".xlsx",sep = ""),
   #encoding = "utf-8")
}

