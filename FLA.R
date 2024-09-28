########################################################################################################
# FLA
########################################################################################################
# Contact: qzhang74@jhu.edu
# Created: 2024/03/07

########################################################################################################
# Initial Set-up
########################################################################################################
# Clear workspace
setwd("~/Desktop/Research2024/FLA/FLA_Meta")

rm(list=ls(all=TRUE))

# Load packages
test<-require(googledrive)   #all gs_XXX() functions for reading data from Google
if (test == FALSE) {
  install.packages("googledrive")
  require(googledrive)
}
test<-require(googlesheets4)   #all gs_XXX() functions for reading data from Google
if (test == FALSE) {
  install.packages("googlesheets4")
  require(googlesheets4)
}
test<-require(plyr)   #rename()
if (test == FALSE) {
  install.packages("plyr")
  require(plyr)
}
test<-require(metafor)   #escalc(); rma();
if (test == FALSE) {
  install.packages("metafor")
  require(metafor)
}
test<-require(robumeta)
if (test == FALSE) {
  install.packages("robumeta")
  require(robumeta)
}
test<-require(weightr) #selection modeling
if (test == FALSE) {
  install.packages("weightr")
  require(weightr)
}
test<-require(mice) #imputation
if (test == FALSE) {
  install.packages("mice")
  require(mice)
}
test<-require(clubSandwich) #coeftest
if (test == FALSE) {
  install.packages("clubSandwich")
  require(clubSandwich)
}
test<-require(tableone)   #CreateTableOne()
if (test == FALSE) {
  install.packages("tableone")
  require(tableone)
}
test<-require(flextable)   
if (test == FALSE) {
  install.packages("flextable")
  require(flextable)
}
test<-require(officer)   
if (test == FALSE) {
  install.packages("officer")
  require(officer)
}
test<-require(tidyverse)   
if (test == FALSE) {
  install.packages("tidyverse")
  require(tidyverse)
}
test<-require(ggrepel)   
if (test == FALSE) {
  install.packages("ggrepel")
  require(ggrepel)
}
test<-require(readxl)   
if (test == FALSE) {
  install.packages("readxl")
  require(readxl)
}

rm(test)

########################################################################################################
# Load data
########################################################################################################
# set up to load from Google
#drive_auth(email = "zhangqiyang0329@gmail.com")
#id <- drive_find(pattern = "FLA_Coding", type = "spreadsheet")$id[1]

# load findings and studies
#gs4_auth(email = "zhangqiyang0329@gmail.com")
# gs4_auth(email = "aneitzel@gmail.com")

findings <- read_excel("FLA_Coding updated 4.2.2024_Final.xlsx", sheet = "Findings")
studies <- read_excel("FLA_Coding updated 4.2.2024_Final.xlsx", sheet = "Studies")

#rm(id)

########################################################################################################
# Clean data
########################################################################################################

# remove irrelevant columns
studies <- subset(studies, select = -c(`First Review`, `Second Review`, `Notes`, `Reason for being dropped`))
findings <- subset(findings, select = -c(`First Review`, `Second Review`, `Notes`))

# merge dataframes
full <- merge(studies, findings, by = c("Study"), all = TRUE, suffixes = c(".s", ".f"))

# rename some cols to shorten
full <- plyr::rename(full, c("Weird (1, 0)" = "Weird"))

# remove any empty rows & columns
full <- subset(full, is.na(full$Drop.f)==TRUE)
full <- subset(full, is.na(full$Drop.s)==TRUE)
full <- subset(full, is.na(full$Study)==FALSE)

# format to correct variable types, "English-relevant Major", "Duration", , 
nums <- c("Sample size", "FemalePercentage", "Age", 
          "Male","Teacher","Treatment.N", 
          "Control.N","Randomized",	"Clustered",
          "T_Mean_Pre", "T_SD_Pre", "C_Mean_Pre", 
          "C_SD_Pre", "T_Mean_Post", "T_SD_Post", 
          "C_Mean_Post", "C_SD_Post", "Effect.Size",
          "Treatment.Cluster", "Control.Cluster",
          "Anxiety Outcome", "Achievement Outcome",
          "Duration.weeks", "SameLanguageFamily", 
          "EnglishL2")
full[nums] <- lapply(full[nums], as.numeric)
rm(nums)

###############################################################
#Create unique identifiers (ES, study, program)
###############################################################
full$ESId <- as.numeric(rownames(full))
full$StudyID <- as.numeric(as.factor(full$Study))

########################################################################################################
# Prep data
########################################################################################################
##### Calculate ESs #####
# calculate pretest ES, SMD is standardized mean difference
full <- escalc(measure = "SMD", m1i = T_Mean_Pre, sd1i = T_SD_Pre, n1i = Treatment.N,
               m2i = C_Mean_Pre, sd2i = C_SD_Pre, n2i = Control.N, data = full)
full$vi <- NULL
full <- plyr::rename(full, c("yi" = "ES_Pre"))

# calculate posttest ES
full <- escalc(measure = "SMD", m1i = T_Mean_Post, sd1i = T_SD_Post, n1i = Treatment.N,
               m2i = C_Mean_Post, sd2i = C_SD_Post, n2i = Control.N, data = full)
full$vi <- NULL
full <- plyr::rename(full, c("yi" = "ES_Post"))

# calculate DID (post ES - pre ES)
full$ES_DID <- full$ES_Post - full$ES_Pre

# put various ES together.  Options:
## 1) used reported ES (so it should be in the Effect.Size column, nothing to do)
## 2) Effect.Size is NA, and DID not missing, replace with that
full$Effect.Size[which(is.na(full$Effect.Size)==TRUE & is.na(full$ES_DID)==FALSE)] <- full$ES_DID[which(is.na(full$Effect.Size)==TRUE & is.na(full$ES_DID)==FALSE)]
## 3) Effect.Size and DID is missing, used adjusted means (from posttest ES col), replace with that
full$Effect.Size[which(is.na(full$Effect.Size)==TRUE & is.na(full$ES_DID)==TRUE)] <- full$ES_Post[which(is.na(full$Effect.Size)==TRUE & is.na(full$ES_DID)==TRUE)]
full$Effect.Size <- as.numeric(full$Effect.Size)

###############################################################
#Calculate meta-analytic variables: Sample sizes
###############################################################
#create full sample/total clusters variables
full$Sample <- full$Treatment.N + full$Control.N
full$Cluster_Total <- full$Treatment.Cluster+full$Control.Cluster

###################################
#Create dummies
###################################
full$EightyPercentFemale <- 0
full$EightyPercentFemale[which(full$`FemalePercentage` >= 0.80)] <- 1

full$FiftyPercentFemale <- 0
full$FiftyPercentFemale[which(full$`FemalePercentage` > 0.50)] <- 1

###############################################################
#Calculate meta-analytic variables: Correct ES for clustering (Hedges, 2007, Eq.19)
###############################################################
#first, create an assumed ICC
full$icc <- NA
full$icc[which(full$Clustered == 1)] <- 0.2

#find average students/cluster
full$Treatment.Cluster.n <- NA
full$Treatment.Cluster.n[which(full$Clustered == 1)] <- round(full$Treatment.N[which(full$Clustered == 1)]/full$Treatment.Cluster[which(full$Clustered == 1)], 0)
full$Control.Cluster.n <- NA
full$Control.Cluster.n[which(full$Clustered == 1)] <- round(full$Control.N[which(full$Clustered == 1)]/full$Control.Cluster[which(full$Clustered == 1)], 0)
# find other parts of equation
full$n.TU <- NA
full$n.TU <- ((full$Treatment.N * full$Treatment.N) - (full$Treatment.Cluster * full$Treatment.Cluster.n * full$Treatment.Cluster.n))/(full$Treatment.N * (full$Treatment.Cluster - 1))
full$n.CU <- NA
full$n.CU <- ((full$Control.N * full$Control.N) - (full$Control.Cluster * full$Control.Cluster.n * full$Control.Cluster.n))/(full$Control.N * (full$Control.Cluster - 1))

# next, calculate adjusted ES, save the originals, then replace only the clustered ones
full$Total.N <- full$`Sample`

full$Effect.Size.adj <- full$Effect.Size * (sqrt(1-full$icc*(((full$Total.N-full$n.TU*full$Treatment.Cluster - full$n.CU*full$Control.Cluster)+full$n.TU + full$n.CU - 2)/(full$Total.N-2))))
full$Effect.Size.adj[full$Effect.Size.adj == "NaN"] <- NA 
View(full[c("Study", "StudyID", "Effect.Size", 
            "Treatment.N", "Control.N", "Effect.Size.adj")])
# save originals, replace for clustered
full$Effect.Size.orig <- full$Effect.Size
full$Effect.Size[which(full$Clustered==1 & is.na(full$Effect.Size.adj)==FALSE)] <- full$Effect.Size.adj[which(full$Clustered==1 & is.na(full$Effect.Size.adj)==FALSE)]

# swaps signs for reverse-coded outcomes (so positive is good and negative is bad)
full$Effect.Size[which(full$ReverseCoded==1)] <- full$Effect.Size[which(full$ReverseCoded==1)]*-1

num <- c("Effect.Size")
full[num] <- lapply(full[num], as.numeric)
rm(num)
################################################################
# Calculate meta-analytic variables: Variances (Lipsey & Wilson, 2000, Eq. 3.23)
################################################################
#calculate standard errors
full$se<-sqrt(((full$Treatment.N+full$Control.N)/(full$Treatment.N*full$Control.N))+((full$Effect.Size*full$Effect.Size)/(2*(full$Treatment.N+full$Control.N))))

#calculate variance
full$var<-full$se*full$se
View(full[c("Study", "StudyID","Effect.Size", 
            "Treatment.N", "Control.N")])

###########################
#forest plot
###########################
#robumeta
forest <- robu(formula = Effect.Size ~ 1, studynum = StudyID, 
               data = full, var.eff.size = var, 
               rho = 0.8, small = TRUE)
print(forest)
forestplot <- forest.robu(forest, es.lab = "Outcome..one.outcome.", 
                          study.lab = "Study",
                          "Effect Size" = Effect.Size)
forestplot

########################################
#meta-regression
########################################
#Null Model
V_list <- impute_covariance_matrix(vi=full$var, cluster=full$StudyID, r=0.8)

MVnull <- rma.mv(yi=Effect.Size,
                 V=V_list,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full,
                 method="REML")
MVnull

# #t-test of each covariate#
MVnull.coef <- coef_test(MVnull, cluster=full$StudyID, vcov="CR2")
MVnull.coef

#Direct intervention type
Direct <- subset(full, full$Direct == 1)
V_list_Direct <- impute_covariance_matrix(vi=Direct$var, cluster=Direct$StudyID, r=0.8)

MVnull_Direct <- rma.mv(yi=Effect.Size,
                 V=V_list_Direct,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=Direct,
                 method="REML")
MVnull_Direct

# #t-test of each covariate#
MVnull.coef_Direct <- coef_test(MVnull_Direct, cluster=Direct$StudyID, vcov="CR2")
MVnull.coef_Direct

#Indirect intervention type
Indirect <- subset(full, full$Indirect == 1)
V_list_Indirect <- impute_covariance_matrix(vi=Indirect$var, cluster=Indirect$StudyID, r=0.8)

MVnull_Indirect <- rma.mv(yi=Effect.Size,
                        V=V_list_Indirect,
                        random=~1 | StudyID/ESId,
                        test="t",
                        data=Indirect,
                        method="REML")
MVnull_Indirect

# #t-test of each covariate#
MVnull.coef_Indirect <- coef_test(MVnull_Indirect, cluster=Indirect$StudyID, vcov="CR2")
MVnull.coef_Indirect
#meta-regression moderators using a fully centered model, "EightyPercentFemale", ,"Culture","FiftyPercentFemale", "Gradelevels"
#remove Age keep Gradelevels, remove EnglishL2 very unbalanced category, "EightyPercentFemale","Age", "FiftyPercentFemale"
#, "Achievement.Outcome", "IF", "Sample.size", "Anxiety", "Achievement"
full$Anxiety <- 0
full$Anxiety[which(full$Anxiety.Outcome==1)] <- 1
full$Achievement <- 0
full$Achievement[which(full$Achievement.Outcome==1)] <- 1
terms <- c("FemalePercentage", "Age",
           "Teacher", "SameLanguageFamily",
           "Duration.weeks", "Weird",
           "Direct")
#formate into formula
#interact <- c("Gradelevels*Teacher") interact
formula <- reformulate(termlabels = c(terms))

#formula <- reformulate(termlabels = c(terms))
#formula

MVfull <- rma.mv(yi=Effect.Size,
                 V=V_list,
                 mods=formula,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full,
                 method="REML")
MVfull

#t-test of each covariate#
MVfull.coef <- coef_test(MVfull, cluster=full$StudyID, vcov="CR2")
MVfull.coef

############################
#Anxiety outcomes only
#############################
full_anxiety <- subset(full, full$Anxiety.Outcome == 1)
V_list_anxiety <- impute_covariance_matrix(vi=full_anxiety$var, cluster=full_anxiety$StudyID, r=0.8)

MVnull_anxiety <- rma.mv(yi=Effect.Size,
                 V=V_list_anxiety,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full_anxiety,
                 method="REML")
MVnull_anxiety

# #t-test of each covariate#
MVnull.coef_anxiety <- coef_test(MVnull_anxiety, cluster=full_anxiety$StudyID, vcov="CR2")
MVnull.coef_anxiety

MVfull_anxiety <- rma.mv(yi=Effect.Size,
                 V=V_list_anxiety,
                 mods=formula,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full_anxiety,
                 method="REML")
MVfull_anxiety

#t-test of each covariate#
MVfull.coef_anxiety <- coef_test(MVfull_anxiety, cluster=full_anxiety$StudyID, vcov="CR2")
MVfull.coef_anxiety

############################
#Achievement outcomes only
#############################
full_achieve <- subset(full, full$Achievement.Outcome == 1)
V_list_achieve <- impute_covariance_matrix(vi=full_achieve$var, cluster=full_achieve$StudyID, r=0.8)

MVnull_achieve <- rma.mv(yi=Effect.Size,
                 V=V_list_achieve,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full_achieve,
                 method="REML")
MVnull_achieve

# #t-test of each covariate#
MVnull.coef_achieve <- coef_test(MVnull_achieve, cluster=full_achieve$StudyID, vcov="CR2")
MVnull.coef_achieve
MVfull_achieve <- rma.mv(yi=Effect.Size,
                 V=V_list_achieve,
                 mods=formula,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full_achieve,
                 method="REML")
MVfull_achieve

#t-test of each covariate#
MVfull.coef_achieve <- coef_test(MVfull_achieve, cluster=full_achieve$StudyID, vcov="CR2")
MVfull.coef_achieve


#################################################################################
#Create Descriptives Table
#################################################################################
# identify variables for descriptive tables (study-level and outcome-level)
vars_study <- c("Weird", "Direct","Gradelevels",
                 "EnglishL2","SameLanguageFamily",
                "Randomized", "Clustered", "Teacher",
                "emotion.targetted", 
                "skills..cognition..targetted ",
                "mindfulness.based", "teamwork.engagement",
                "technology.assisted")

vars_outcome <- c("Anxiety.Outcome", "Achievement.Outcome")

# To make this work, you will need a df that is at the study-level for study-level 
# variables (such as research design) you may have already created this (see above, with study-level ESs), but if you didn't, here is an easy way:
# 1) make df with *only* the study-level variables of interest and studyIDs in it
study_level_full <- full[c("StudyID", "Weird", "Direct", "Gradelevels",
                           "EnglishL2","SameLanguageFamily",
                           "Randomized", "Clustered", "Teacher",
                           "emotion.targetted", 
                           "skills..cognition..targetted",
                           "mindfulness.based", "teamwork.engagement",
                           "technology.assisted")]


outcome_level_full <- full[c("Anxiety.Outcome", "Achievement.Outcome")]

# 2) remove duplicated rows
study_level_full <- unique(study_level_full)
# 3) make sure it is the correct number of rows (should be same number of studies you have)
length(study_level_full$StudyID)==length(unique(study_level_full$StudyID))
# don't skip step 3 - depending on your data structure, some moderators can be
# study-level in one review, but outcome-level in another

# create the table "chunks"
table_study_df <- as.data.frame(print(CreateTableOne(vars = vars_study, data = study_level_full, 
                                                     includeNA = TRUE, 
                                                     factorVars = c("Weird", "Direct","Gradelevels",
                                                                    "EnglishL2","SameLanguageFamily",
                                                                    "Randomized", "Clustered", "Teacher",
                                                                    "emotion.targetted", 
                                                                    "skills..cognition..targetted ",
                                                                    "mindfulness.based", "teamwork.engagement",
                                                                    "technology.assisted")), 
                                      showAllLevels = TRUE))
table_outcome_df <- as.data.frame(print(CreateTableOne(vars = vars_outcome, data = full, includeNA = TRUE,
                                                       factorVars = c("Anxiety.Outcome", "Achievement.Outcome")), 
                                        showAllLevels = TRUE))
rm(vars_study, vars_outcome)


################################
# Descriptives Table Formatting
################################
table_study_df$Category <- row.names(table_study_df)
rownames(table_study_df) <- c()
table_study_df <- table_study_df[c("Category", "level", "Overall")]
table_study_df$Category[which(substr(table_study_df$Category, 1, 1)=="X")] <- NA
table_study_df$Category <- gsub(pattern = "\\..mean..SD..", replacement = "", x = table_study_df$Category)
table_study_df$Overall <- gsub(pattern = "\\( ", replacement = "\\(", x = table_study_df$Overall)
table_study_df$Overall <- gsub(pattern = "\\) ", replacement = "\\)", x = table_study_df$Overall)
table_study_df$Category <- gsub(pattern = "\\.", replacement = "", x = table_study_df$Category)
table_study_df$Category[which(table_study_df$Category=="n")] <- "Total Studies"
table_study_df$level[which(table_study_df$level=="1")] <- "Yes"
table_study_df$level[which(table_study_df$level=="0")] <- "No                                                                              "
# fill in blank columns (to improve merged cells later)
for(i in 1:length(table_study_df$Category)) {
  if(is.na(table_study_df$Category[i])) {
    table_study_df$Category[i] <- table_study_df$Category[i-1]
  }
}

table_outcome_df$Category <- row.names(table_outcome_df)
rownames(table_outcome_df) <- c()
table_outcome_df <- table_outcome_df[c("Category", "level", "Overall")]
table_outcome_df$Category[which(substr(table_outcome_df$Category, 1, 1)=="X")] <- NA
table_outcome_df$Overall <- gsub(pattern = "\\( ", replacement = "\\(", x = table_outcome_df$Overall)
table_outcome_df$Overall <- gsub(pattern = "\\) ", replacement = "\\)", x = table_outcome_df$Overall)
table_outcome_df$Category <- gsub(pattern = "\\.", replacement = "", x = table_outcome_df$Category)
table_outcome_df$Category[which(table_outcome_df$Category=="n")] <- "Total Effect Sizes"
# fill in blank columns (to improve merged cells later)
for(i in 1:length(table_outcome_df$Category)) {
  if(is.na(table_outcome_df$Category[i])) {
    table_outcome_df$Category[i] <- table_outcome_df$Category[i-1]
  }
}

########################
#Output officer
########################
myreport<-read_docx()
# Descriptive Table
myreport <- body_add_par(x = myreport, value = "Table 4: Descriptive Statistics", style = "Normal")
descriptives_study <- flextable(head(table_study_df, n=nrow(table_study_df)))
descriptives_study <- add_header_lines(descriptives_study, values = c("Study Level"), top = FALSE)
descriptives_study <- theme_vanilla(descriptives_study)
descriptives_study <- merge_v(descriptives_study, j = c("Category"))
myreport <- body_add_flextable(x = myreport, descriptives_study)

descriptives_outcome <- flextable(head(table_outcome_df, n=nrow(table_outcome_df)))
descriptives_outcome <- delete_part(descriptives_outcome, part = "header")
descriptives_outcome <- add_header_lines(descriptives_outcome, values = c("Outcome Level"))
descriptives_outcome <- theme_vanilla(descriptives_outcome)
descriptives_outcome <- merge_v(descriptives_outcome, j = c("Category"))
myreport <- body_add_flextable(x = myreport, descriptives_outcome)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")

########################
# MetaRegression Table 1
########################
MVnull.coef
str(MVnull.coef)
MVnull.coef$coef <- row.names(as.data.frame(MVnull.coef))
row.names(MVnull.coef) <- c()
MVnull.coef <- MVnull.coef[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]
MVnull.coef
str(MVnull.coef)

MVfull.coef$coef <- row.names(as.data.frame(MVfull.coef))
row.names(MVfull.coef) <- c()
MVfull.coef <- MVfull.coef[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]

# MetaRegression Table
model_null <- flextable(head(MVnull.coef, n=nrow(MVnull.coef)))
colkeys <- c("beta", "SE", "tstat", "df_Satt")
model_null <- colformat_double(model_null,  j = colkeys, digits = 2)
model_null <- colformat_double(model_null,  j = c("p_Satt"), digits = 3)
#model_null <- autofit(model_null)
model_null <- add_header_lines(model_null, values = c("Null Model"), top = FALSE)
model_null <- theme_vanilla(model_null)

myreport <- body_add_par(x = myreport, value = "Table 5: Model Results", style = "Normal")
myreport <- body_add_flextable(x = myreport, model_null)

# MetaRegression Table
model_full <- flextable(head(MVfull.coef, n=nrow(MVfull.coef)))
model_full <- colformat_double(model_full,  j = c("beta"), digits = 2)
model_full <- colformat_double(model_full,  j = c("SE"), digits = 3)
model_full <- colformat_double(model_full,  j = c("tstat"), digits = 2)
model_full <- colformat_double(model_full,  j = c("df_Satt"), digits = 2)
model_full <- colformat_double(model_full,  j = c("p_Satt"), digits = 3)
#model_full <- autofit(model_full)
model_full <- delete_part(model_full, part = "header")
model_full <- add_header_lines(model_full, values = c("Meta-Regression"))
model_full <- theme_vanilla(model_full)

myreport <- body_add_flextable(x = myreport, model_full)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")

########################
# MetaRegression Table 2
########################
MVnull.coef_anxiety
str(MVnull.coef_anxiety)
MVnull.coef_anxiety$coef <- row.names(as.data.frame(MVnull.coef_anxiety))
row.names(MVnull.coef_anxiety) <- c()
MVnull.coef_anxiety <- MVnull.coef_anxiety[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]
MVnull.coef_anxiety
str(MVnull.coef_anxiety)

MVfull.coef_anxiety$coef <- row.names(as.data.frame(MVfull.coef_anxiety))
row.names(MVfull.coef_anxiety) <- c()
MVfull.coef_anxiety <- MVfull.coef_anxiety[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]

# MetaRegression Table
model_null_anxiety <- flextable(head(MVnull.coef_anxiety, n=nrow(MVnull.coef_anxiety)))
colkeys <- c("beta", "SE", "tstat", "df_Satt")
model_null_anxiety <- colformat_double(model_null_anxiety,  j = colkeys, digits = 2)
model_null_anxiety <- colformat_double(model_null_anxiety,  j = c("p_Satt"), digits = 3)
#model_null <- autofit(model_null)
model_null_anxiety <- add_header_lines(model_null_anxiety, values = c("Null Model"), top = FALSE)
model_null_anxiety <- theme_vanilla(model_null_anxiety)

myreport <- body_add_par(x = myreport, value = "Table 6: Model Results for Anxiety Outcomes", style = "Normal")
myreport <- body_add_flextable(x = myreport, model_null_anxiety)

# MetaRegression Table
model_full_anxiety <- flextable(head(MVfull.coef_anxiety, n=nrow(MVfull.coef_anxiety)))
model_full_anxiety <- colformat_double(model_full_anxiety,  j = c("beta"), digits = 2)
model_full_anxiety <- colformat_double(model_full_anxiety,  j = c("SE"), digits = 3)
model_full_anxiety <- colformat_double(model_full_anxiety,  j = c("tstat"), digits = 2)
model_full_anxiety <- colformat_double(model_full_anxiety,  j = c("df_Satt"), digits = 2)
model_full_anxiety <- colformat_double(model_full_anxiety,  j = c("p_Satt"), digits = 3)
#model_full <- autofit(model_full)
model_full_anxiety <- delete_part(model_full_anxiety, part = "header")
model_full_anxiety <- add_header_lines(model_full_anxiety, values = c("Meta-Regression"))
model_full_anxiety <- theme_vanilla(model_full_anxiety)

myreport <- body_add_flextable(x = myreport, model_full_anxiety)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")

########################
# MetaRegression Table 3
########################
MVnull.coef_achieve
str(MVnull.coef_achieve)
MVnull.coef_achieve$coef <- row.names(as.data.frame(MVnull.coef_achieve))
row.names(MVnull.coef_achieve) <- c()
MVnull.coef_achieve <- MVnull.coef_achieve[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]
MVnull.coef_achieve
str(MVnull.coef_achieve)

MVfull.coef_achieve$coef <- row.names(as.data.frame(MVfull.coef_achieve))
row.names(MVfull.coef_achieve) <- c()
MVfull.coef_achieve <- MVfull.coef_achieve[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]

# MetaRegression Table
model_null_achieve <- flextable(head(MVnull.coef_achieve, n=nrow(MVnull.coef_achieve)))
colkeys <- c("beta", "SE", "tstat", "df_Satt")
model_null_achieve <- colformat_double(model_null_achieve,  j = colkeys, digits = 2)
model_null_achieve <- colformat_double(model_null_achieve,  j = c("p_Satt"), digits = 3)
#model_null <- autofit(model_null)
model_null_achieve <- add_header_lines(model_null_achieve, values = c("Null Model"), top = FALSE)
model_null_achieve <- theme_vanilla(model_null_achieve)

myreport <- body_add_par(x = myreport, value = "Table 7: Model Results for Achievement Outcomes", style = "Normal")
myreport <- body_add_flextable(x = myreport, model_null_achieve)

# MetaRegression Table
model_full_achieve <- flextable(head(MVfull.coef_achieve, n=nrow(MVfull.coef_anxiety)))
model_full_achieve <- colformat_double(model_full_achieve,  j = c("beta"), digits = 2)
model_full_achieve <- colformat_double(model_full_achieve,  j = c("SE"), digits = 3)
model_full_achieve <- colformat_double(model_full_achieve,  j = c("tstat"), digits = 2)
model_full_achieve <- colformat_double(model_full_achieve,  j = c("df_Satt"), digits = 2)
model_full_achieve <- colformat_double(model_full_achieve,  j = c("p_Satt"), digits = 3)
#model_full <- autofit(model_full)
model_full_achieve <- delete_part(model_full_achieve, part = "header")
model_full_achieve <- add_header_lines(model_full_achieve, values = c("Meta-Regression"))
model_full_achieve <- theme_vanilla(model_full_achieve)

myreport <- body_add_flextable(x = myreport, model_full_achieve)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")
# # Marginal Means Table
# marginalmeans <- flextable(head(means, n=nrow(means)))
# colkeys <- c("moderator", "group", "beta", "SE", "tstat", "df_Satt")
# marginalmeans <- colformat_double(marginalmeans,  j = colkeys, digits = 2)
# marginalmeans <- colformat_double(marginalmeans,  j = c("p_Satt"), digits = 3)
# rm(colkeys)
# marginalmeans <- theme_vanilla(marginalmeans)
# marginalmeans <- merge_v(marginalmeans, j = c("moderator"))
# myreport <- body_add_par(x = myreport, value = "Table: Marginal Means", style = "Normal")
# myreport <- body_add_flextable(x = myreport, marginalmeans)

# Write to word doc
file = paste("TableResults.docx", sep = "")
print(myreport, file)

#publication bias , selection modeling
full_y <- full$Effect.Size
full_v <- full$var
weightfunct(full_y, full_v)
weightfunct(full_y, full_v, steps = c(.025, .50, .95, 1))
