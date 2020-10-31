#### Part A: 讀取套件 ####
if(!require(xlsx)) {install.packages("xlsx"); library(xlsx)}
if(!require(survival)) {install.packages("survival"); library(survival)}
if(!require(lme4)) {install.packages("lme4"); library(lme4)}
if(!require(nlme)) {install.packages("nlme"); library(nlme)}
if(!require(psych)) {install.packages("psych"); library(psych)}
if(!require(cccrm)) {install.packages("cccrm"); library(cccrm)} # 做信度分析
library(readxl)
#### Part A: 清除物件 & 設定Output匯出位置與Input資料位置 & 讀取資料 ####
rm(list = ls())
# Output匯出位置
setwd("./data")

# Output匯出名稱
save.name = 'Output_2018.10.14.xlsx'
# 標計fu的sheet需要讀入幾個row (不設定會讀到後面多餘的空白欄位)
L=339
# Input資料位置
name = paste0("Autotransplantation survey data (coding) 20181009.xlsx")

# 分別讀取4個sheets的資料

Basic <- as.data.frame(read_xlsx(name,sheet= 1)[, 1:7])
Pre = as.data.frame(read_xlsx(name, sheet= 2))
Trt = as.data.frame(read_xlsx(name, sheet= 3))
fu <- as.data.frame(read_xlsx(name, sheet= 4))[-339,]

#### Part B: 合併/整理/重新coding 4個sheets的資料 ####
reset = function(){
  ##### Basic information #####
  # Basic information中需要的變數
  visit.ym = as.numeric(as.character(Basic[,1]))
  chart.no = as.numeric(as.character(Basic[, "Chart No,"]))
  No = 1:96 # 原使資料有96人，之後會排除一些人
  name = as.character(Basic$Name)
  birth = Basic$Birthday
  sex = Basic$Sex
  ##### Pre-op status #####
  # Pre-op status(病人治療前狀態)中需要的變數
  donor.tooth.position = Pre[,'Donor tooth position']
  donor.tooth.shape = Pre[,'Root shape']
  donor.tooth.impaction = as.character(Pre[,"Impaction classification"])
  donor.tooth.h = rep(0,length(donor.tooth.impaction))
  donor.tooth.v = rep(0,length(donor.tooth.impaction))
  donor.tooth.hv = rep(0,length(donor.tooth.impaction))
  for(i in 1:length(donor.tooth.impaction)){  # 分別將垂直/水平原本的文字轉成數字
    if(grepl('III',donor.tooth.impaction[i])){
      donor.tooth.h[i] = 3
    }else if(grepl('II',donor.tooth.impaction[i])){
      donor.tooth.h[i] = 2
    }else if(grepl('I',donor.tooth.impaction[i])){
      donor.tooth.h[i] = 1
    }
    if(grepl('C',donor.tooth.impaction[i])){
      donor.tooth.v[i] = 3
    }else if(grepl('B',donor.tooth.impaction[i])){
      donor.tooth.v[i] = 2
    }else if(grepl('A',donor.tooth.impaction[i]) & donor.tooth.impaction[i]!='NA'){
      donor.tooth.v[i] = 1
    }
    # 將垂直和水平的合併成一個新的變數
    if(donor.tooth.h[i] == 0 & donor.tooth.v[i]==0) donor.tooth.hv[i]=0
    else if(donor.tooth.h[i] == 1 & donor.tooth.v[i]==0) donor.tooth.hv[i]=1
    else if(donor.tooth.h[i] == 2 & donor.tooth.v[i]==0) donor.tooth.hv[i]=2
    else if(donor.tooth.h[i] == 3 & donor.tooth.v[i]==0) donor.tooth.hv[i]=3
    else if(donor.tooth.h[i] == 0 & donor.tooth.v[i]==1) donor.tooth.hv[i]=4
    else if(donor.tooth.h[i] == 1 & donor.tooth.v[i]==1) donor.tooth.hv[i]=5
    else if(donor.tooth.h[i] == 2 & donor.tooth.v[i]==1) donor.tooth.hv[i]=6
    else if(donor.tooth.h[i] == 3 & donor.tooth.v[i]==1) donor.tooth.hv[i]=7
    else if(donor.tooth.h[i] == 0 & donor.tooth.v[i]==2) donor.tooth.hv[i]=8
    else if(donor.tooth.h[i] == 1 & donor.tooth.v[i]==2) donor.tooth.hv[i]=9
    else if(donor.tooth.h[i] == 2 & donor.tooth.v[i]==2) donor.tooth.hv[i]=10
    else if(donor.tooth.h[i] == 3 & donor.tooth.v[i]==2) donor.tooth.hv[i]=11
    else if(donor.tooth.h[i] == 0 & donor.tooth.v[i]==3) donor.tooth.hv[i]=12
    else if(donor.tooth.h[i] == 1 & donor.tooth.v[i]==3) donor.tooth.hv[i]=13
    else if(donor.tooth.h[i] == 2 & donor.tooth.v[i]==3) donor.tooth.hv[i]=14
    else if(donor.tooth.h[i] == 3 & donor.tooth.v[i]==3) donor.tooth.hv[i]=15
  }
  donor.tooth.h = as.factor(donor.tooth.h)
  donor.tooth.v = as.factor(donor.tooth.v)
  donor.tooth.hv = as.factor(donor.tooth.hv)
  
  root.formation.stage = Pre[,'Root formation stage']
  recipient.tooth.position = Pre[,"Recipient tooth position"]
  drill = as.numeric(as.character(Pre[,"Extraction socket or Drill"]))
  donor.tooth.dx = Pre[,"periodontitis or not (新)"]  # 為空缺牙齒缺失的原因，請用新整理過的此變數
  # 另外一個類似的變數是比較細節的原始變數
  caries = rep(NA,96)# Pre[,'Caries']         # 經討論後，因不同類別比例相差過於懸殊而未使用的變數
  ##### Treatment related #####
  # Treatment related(病人治療方式/過程記錄)中需要的變數
  fixation.type = NA# Trt[,'Fixation.type']   # 經討論後，因不同類別比例相差過於懸殊而未使用的變數
  dressing.time = as.character(Trt[,'Dressing time'])
  dressing.day = rep(NA,length(dressing.time))
  for(i in 1:length(dressing.time)){
    temp1 = as.numeric(substr(dressing.time[i],1,2))
    temp2 = substr(dressing.time[i],3,15)
    if(grepl('week',temp2)){
      dressing.day[i] = temp1*7
    }else if(grepl('month',temp2)){
      dressing.day[i] = temp1*30
    }else{
      dressing.day[i] = temp1
    }
  }
  severe.curvature = as.numeric(as.character(Trt[,'severe curvature'])=='+')
  Cshape = as.numeric(as.character(Trt[,'c-shaped'])=='+')
  debridement.day = Trt[,'Debridement time']
  doctor = as.numeric(as.character(Trt[,'根管治療醫師'])=='B')
  
  ##### fu #####
  # 每次病人回診的outcome需要的變數
  op.date = birth; 
  count = 1
  opened.date = birth
  # 第5(Probing.depth.coding)和第6(Bone.heal.coding)個變數是判斷手術是否成功的兩個指標
  # 其餘為病人編號、初診/手術/回診日期
  sub_fu = fu[,c("Chart No-","OP  date","Opened date","Follow Date","Probing depth coding","Bone heal coding")]
  sub_fu[,1] = as.numeric(as.character(sub_fu[,1])); 
  index = 1;
  # 把原本標記為NA的通通換成missing value
  # 記得原始檔案有的NA是全形字，有的是半形字，請修改原本檔案後再丟進來，或是下面指令修改為也能抓全形字
  for(i in 5:6){  
    sub_fu[,i] = as.character(sub_fu[,i])
    sub_fu[which(sub_fu[,i] == 'NA'),i] = NA
  }
  
  ##### Success #####
  # 手術剛開始都先當成失敗，直到囊袋深度與骨頭癒合程度達到標準，才視為手術成功
  # 由於囊袋深度(Probing.depth.coding)遺失值比較多，因此分兩種來看
  # 第一種：骨頭癒合狀況不明的追蹤資料視為遺失值
  #         囊袋深度達到治癒的判斷標準或遺失，且骨頭完全癒合的情況視為治療成功
  success.with.na = as.numeric((sub_fu[,5] == 1|is.na(sub_fu[,5])) & sub_fu[,6] == 3)
  # 第二種：囊袋深度未紀錄/骨頭癒合狀況不明的追蹤資料都視為遺失值
  #         囊袋深度達到治癒的判斷標準，且骨頭完全癒合的情況視為治療成功
  success = as.numeric(sub_fu[,5] == 1 & sub_fu[,6] == 3)
  colnames(sub_fu)[1:4] = c('chart no','op date','opened date','follow date')
  
  ##### No #####
  # 重新給病人編號，便於對照/修改/查閱
  sub_fu = cbind(No = NA, sub_fu)
  count <- 1
  while(count<97){
    if(chart.no[count]==sub_fu[index,2])  {
      op.date[count] = sub_fu[index,3]
      opened.date[count] = sub_fu[index,4]
      sub_fu[index,'No'] = count
      if(chart.no[count]!=6105631){ # 這個病人的編號在fu的sheet一直都有問題，因此特別修改
        while(chart.no[count]==sub_fu[index,2] & index<=338){
          index = index + 1
          if(index<=338){
            sub_fu[index,'No'] = count
          }
        };count = count + 1
      }else{
        index = index + 1
        count = count + 1
        sub_fu[index,'No'] = count
        op.date[count] = sub_fu[index,3]
        opened.date[count] = sub_fu[index,4]
        index = index + 1
        count = count + 1
      }
    }
    print(c(count, index))
    if(count == 97 || index ==339) break;
  }
  
  # 使用電腦計算追蹤天數 & 合併手術成功與否
  sub_fu = cbind(sub_fu,follow.day = as.numeric(sub_fu[,'follow date'] - sub_fu[,'op date']), 
                 success, success.with.na)
  # 合併除了fu以外的所有資料
  base = data.frame(No, name,visit.ym,chart.no,birth,age = floor(as.numeric((op.date - birth)/365.25)),
                    sex,donor.tooth.position,root.formation.stage, impacted = as.numeric(donor.tooth.hv!=0),
                    donor.tooth.shape,donor.tooth.h,donor.tooth.v, donor.tooth.hv,donor.caries=caries, 
                    severe.curvature, Cshape, recipient.tooth.position,drill,donor.tooth.dx, doctor,
                    fixation.type = as.factor(fixation.type), op.date,opened.date, dressing.day,
                    opened.day = as.numeric(opened.date - op.date)/86400, debridement.day)
  base = cbind(base, endodontics.day = dressing.day + debridement.day)
  # 特別說要看這兩個連續變數按照14天與30天做切割的結果
  opened14 = rep(0, length(base[,'opened.day']))
  dressing30 = opened14
  debridement30 = opened14
  agec = opened14
  upper = opened14
  opened14[which(base[,'opened.day']>14)] = 1
  dressing30[which(base[,'dressing.day']>30)] = 1
  debridement30[which(base[,'debridement.day']>30)] = 1
  # 年齡分三組：20~30為0；30~40為1；40以上為2
  agec[base[,'age']>=30 & base[,'age']<40] = 1
  agec[base[,'age']>=40] = 2
  # 標記是否為上排牙齒
  upper[base[,'donor.tooth.position']<30] = 1
  # 再次合併資料(尚未合併fu的資料)
  base = cbind(base, opened14, dressing30, debridement30, agec = as.factor(agec), upper)
  # cL = ncol(base) # 用來看變數名稱而已
  # 重新排列欄位順序
  base = base[,c('No', 'name', 'visit.ym', 'chart.no', 'birth', 'age', 'agec', 'sex', 'doctor',
                 'donor.tooth.position', 'upper', 'root.formation.stage',
                 'donor.tooth.shape', 'donor.tooth.h', 'donor.tooth.v', 'donor.tooth.hv', 'impacted',
                 'severe.curvature', 'Cshape', 'recipient.tooth.position', 
                 'donor.tooth.dx', 'op.date', 'opened.date', 
                 'dressing.day', 'opened.day', 'debridement.day', 'endodontics.day', 
                 'opened14', 'dressing30')]
  # 以上合併為長形資料，因此此變數紀錄為第幾次追蹤的結果
  times = as.vector(table(sub_fu[,'No']))
  temp = as.data.frame(matrix(, 0, ncol(base)))
  for(i in 1:length(times)){
    count = times[i]
    while(count>0){
      temp = rbind(temp,base[i,])
      count = count - 1
    }
  }
  ##### Output(Original data) #####
  # 將fu的資料併入
  long.data = cbind(temp, sub_fu[,5:10])
  # 排掉年齡未滿20的病人
  # 排掉後來決定不使用的兩個變數
  long.data = long.data[long.data[,'age']>=20,
                        which(colnames(long.data)!='donor.caries' & colnames(long.data)!='fixation.type')]
  
  # 匯出整理過後的長形資料(4個sheets整理成1個sheet)
  # write.xlsx(long.data,save.name, sheetName='full', append=T, row.names=F, showNA=F)
  
  # 清掉讀入的資料，只保留整理過的資料
  rm(list = ls()[which(ls()!= 'long.data' & ls()!= 'Basic' & ls()!= 'Pre' & ls()!= 'Trt' 
                       & ls()!= 'fu' & ls()!= 'L' & ls()!= 'save.name')])
  return(long.data)
}; long.data = reset()

write.csv(long.data, "long.data.csv", row.names=F)

#### Part C: 輸出整理好的資料(2 outcome) #####
# data：放入的資料
# event：為outcome是哪個變數，
#        填入success(囊袋深度和骨頭癒合程度都看)或Bone.heal.coding(只看骨頭癒合程度)
# na：是否要排掉missing值


sur.data = function(data, event, na){
  if(event != 'Bone heal coding'){
    if(na == F){
      data = data[!is.na(data[,event]),]
      temp=unique(data[,'No'])
    }else{
      temp=unique(data[,'No'])
    }; n = 96
    index = rep(NA,n)
    for(i in temp){
      tag = which(data[,'No']==i)
      tag.s = which(c(1,data[tag,event]) == 1)-1
      if(max(tag.s)==0){
        index[i] = max(tag)
      }else{
        index[i] = tag[min(tag.s[-1])]
      }
    }
    index = index[!is.na(index)]
    if(event == 'success'){
      data = data[index,-ncol(data)]
    }else{
      data = data[index,-(ncol(data)-1)]
    }
    
  }else{
    if(na == F){
      data = data[!is.na(data[,event]),]
      temp=unique(data[,'No'])
    }else{
      temp=unique(data[,'No'])
    }; n = 96
    index = rep(NA,n)
    for(i in temp){
      tag = which(data[,'No']==i)
      tag.s = which(c(3,data[tag,event]) == 3)-1
      if(max(tag.s)==0){
        index[i] = max(tag)
      }else{
        index[i] = tag[min(tag.s[-1])]
      }
    }
    index = index[!is.na(index)]
    data = data[index,-(ncol(data)-0:1)]
    
    complete = rep(NA,nrow(data))
    complete[data[,event]==3] = 1
    complete[data[,event]==1 | data[,event]==2] = 0
    data = cbind(data, complete)
  }
  return(data)
}
s.data = sur.data(long.data,'success',F)
c.data = sur.data(long.data,'Bone heal coding',F)
#### Part D: 畫KM curve ####
require("survival")
require("survminer")
names(c.data)
fit<- survfit(Surv(follow.day, success.with.na) ~ severe.curvature, data = c.data)
ggsurv <- ggsurvplot(fit,
                     pval = TRUE, conf.int = T,
                     risk.table = TRUE,
                     surv.median.line = "hv",# Add risk table
                     risk.table.col = "strata", # Change risk table color by groups
                     ggtheme = theme_bw(), # Change ggplot2 theme
                     palette = c("#E7B800", "#2E9FDF") #general
                     # palette = c("#0099CC","#FF6666") #sex
                     # palette = c("#BF7750", "#357E68", "#3A5A7D") #age
)
print(ggsurv)

# Drawing a horizontal line at 50% survival
surv_median <- as.vector(summary(fit)$table[, "median"])
df <- data.frame(x1 = surv_median, x2 = surv_median,
                 y1 = rep(0, length(surv_median)), y2 = rep(0.5, length(surv_median)))

ggsurv$plot <- ggsurv$plot + 
  geom_segment(aes(x = 0, y = 0.5, xend = max(surv_median), yend = 0.5),
               linetype = "dashed", size = 0.5)+ # horizontal segment
  geom_segment(aes(x = x1, y = y1, xend = x2, yend = y2), data = df,
               linetype = "dashed", size = 0.5) # vertical segments

print(ggsurv)
