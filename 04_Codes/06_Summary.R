# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MSD CHC 2020Q1Q2
# Purpose:      Summary
# programmer:   Zhe Liu
# Date:         2020-08-25
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


##---- Readin data ----
# pack info
pack.ref <- fread("02_Inputs/cn_prod_ref_201912_1.txt", 
                  stringsAsFactors = FALSE, sep = "|") %>% 
  distinct() %>% 
  mutate(Pack_Id = stri_pad_left(Pack_Id, 7, 0))

# corp info
corp.ref <- fread("02_Inputs/cn_corp_ref_201912_1.txt", 
                  stringsAsFactors = FALSE, sep = "|") %>% 
  distinct()

# pack & corp
corp.pack <- pack.ref %>% 
  distinct(Pack_Id, Prd_desc, Pck_Desc, Corp_ID, PckSize_Desc) %>% 
  left_join(corp.ref, by = "Corp_ID") %>% 
  select(packid = Pack_Id, Pck_Desc, Corp_Desc, PckSize_Desc)

# product desc
prod.desc <- read.csv("02_Inputs/pfc与ims数据对应_20200824.csv") %>% 
  mutate(Pack_Id = stri_pad_left(Pack_Id, 7, 0)) %>% 
  select(packid = Pack_Id, Prd_desc = `商品名`)


##---- Result ----
# Servier result
servier.result <- read.xlsx("02_Inputs/08_Servier_CHC_2016Q4_2020Q2.xlsx") %>% 
  filter(Channel == 'CHC', 
         MKT == 'OAD', 
         Date %in% c('2020Q1', '2020Q2'), 
         Molecule_Desc %in% market.def$Molecule) %>% 
  mutate(Pack_ID = stri_pad_left(Pack_ID, 7, 0)) %>% 
  select(Pack_ID, Channel, Province, City, Date, ATC3, MKT, Molecule_Desc, 
         Prod_Desc, Pck_Desc, Corp_Desc, Sales, Units, DosageUnits)

# final result
msd.city.result <- msd.price %>% 
  filter(!(city %in% servier.target.city)) %>% 
  left_join(corp.pack, by = "packid") %>% 
  left_join(prod.desc, by = "packid") %>% 
  mutate(Channel = "CHC",
         dosageunits = PckSize_Desc * units) %>% 
  group_by(Pack_ID = packid, Channel, Province = province, City = city, 
           Date = quarter, ATC3 = atc3, MKT = market, Molecule_Desc = molecule, 
           Prod_Desc = Prd_desc, Pck_Desc, Corp_Desc) %>% 
  summarise(Sales = sum(sales, na.rm = TRUE),
            Units = sum(units, na.rm = TRUE),
            DosageUnits = sum(dosageunits, na.rm = TRUE)) %>% 
  ungroup() %>% 
  bind_rows(servier.result) %>% 
  mutate(Pack_ID = if_else(stri_sub(Pack_ID, 1, 5) == '06470', 
                           stri_paste('64895', stri_sub(Pack_ID, 6, 7)), 
                           Pack_ID), 
         Pack_ID = if_else(stri_sub(Pack_ID, 1, 5) == '47775', 
                           stri_paste('58906', stri_sub(Pack_ID, 6, 7)), 
                           Pack_ID), 
         Corp_Desc = if_else(Prod_Desc == "GLUCOPHAGE", "MERCK GROUP", Corp_Desc),
         Corp_Desc = if_else(Prod_Desc == 'ONGLYZA', 'ASTRAZENECA GROUP', Corp_Desc), 
         Corp_Desc = if_else(Corp_Desc == "LVYE GROUP", "LUYE GROUP", Corp_Desc), 
         # price4 = 103.26608, 
         # price2 = 52.12415, 
         Units = if_else(City == '上海' & Pack_ID == '4268604', 
                         Sales / 103.26608, 
                         Units), 
         DosageUnits = if_else(City == '上海' & Pack_ID == '4268604', 
                               Units * 14, 
                               DosageUnits), 
         Units = if_else(City == '上海' & Pack_ID == '4268602', 
                         Sales / 52.12415, 
                         Units), 
         DosageUnits = if_else(City == '上海' & Pack_ID == '4268602', 
                               Units * 7, 
                               DosageUnits)) %>% 
  group_by(Pack_ID, Channel, Province, City, Date, ATC3, MKT, Molecule_Desc, 
           Prod_Desc, Pck_Desc, Corp_Desc) %>% 
  summarise(Sales = sum(Sales, na.rm = TRUE),
            Units = sum(Units, na.rm = TRUE),
            DosageUnits = sum(DosageUnits, na.rm = TRUE)) %>% 
  ungroup() %>% 
  filter(Sales > 0, Units > 0, DosageUnits > 0) %>% 
  arrange(Date, Province, City, Pack_ID)

msd.nation.result <- msd.city.result %>% 
  group_by(Pack_ID, Channel, Province = "National", City = "National", 
           Date, ATC3, MKT, Molecule_Desc, Prod_Desc, Pck_Desc, Corp_Desc) %>% 
  summarise(Sales = sum(Sales, na.rm = TRUE),
            Units = sum(Units, na.rm = TRUE),
            DosageUnits = sum(DosageUnits, na.rm = TRUE)) %>% 
  ungroup()

# adj.factor <- read.xlsx("02_Inputs/Adjust_Factor.xlsx") %>% 
#   mutate(Pack_ID = stri_pad_left(Pack_ID, 7, 0)) %>% 
#   setDT() %>% 
#   melt(id.vars = c("Prod_Desc", "Pack_ID"), 
#        variable.name = "City", 
#        value.name = "factor", 
#        variable.factor = FALSE)

msd.result <- msd.city.result %>% 
  filter(City %in% target.city) %>% 
  bind_rows(msd.nation.result) %>% 
  group_by(Pack_ID, Channel, Province, City, Date, ATC3, MKT, Molecule_Desc, 
           Prod_Desc, Pck_Desc, Corp_Desc) %>% 
  summarise(Sales = sum(Sales, na.rm = TRUE),
            Units = sum(Units, na.rm = TRUE),
            DosageUnits = sum(DosageUnits, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(Sales = case_when(
           Pack_ID == '4268604' & City == '福州' ~ Sales * 1.8, 
           Pack_ID == '4268604' & City == '广州' ~ Sales * 0.7, 
           Pack_ID == '4268602' & City == '广州' ~ Sales * 4, 
           Pack_ID == '4268604' & City == '上海' ~ Sales * 1700 * 0.8, 
           Pack_ID == '4268602' & City == '上海' ~ Sales * 4100 * 0.8, 
           Pack_ID == '4268602' & City == '苏州' ~ Sales * 1.2, 
           Pack_ID == '4268604' & City == 'National' ~ Sales * 1.08, 
           Pack_ID == '4268602' & City == 'National' ~ Sales * 1.6 * 1.08, 
           Prod_Desc == 'JANUVIA' & City == '南京' ~ Sales * 0.8, 
           Prod_Desc == 'JANUMET' & City == 'National' ~ Sales * 3, 
           Prod_Desc == 'DIAMICRON' & City %in% c('杭州', '南京', '上海', '苏州') ~ Sales * 2, 
           Prod_Desc == 'DIAMICRON' & City == 'National' ~ Sales * 1.1, 
           Prod_Desc == 'VICTOZA' & City == 'National' ~ Sales * 2.5, 
           TRUE ~ Sales
         ), 
         Units = case_when(
           Pack_ID == '4268604' & City == '福州' ~ Units * 1.8, 
           Pack_ID == '4268604' & City == '广州' ~ Units * 0.7, 
           Pack_ID == '4268602' & City == '广州' ~ Units * 4, 
           Pack_ID == '4268604' & City == '上海' ~ Units * 1700 * 0.8, 
           Pack_ID == '4268602' & City == '上海' ~ Units * 4100 * 0.8, 
           Pack_ID == '4268602' & City == '苏州' ~ Units * 1.2, 
           Pack_ID == '4268604' & City == 'National' ~ Units * 1.08, 
           Pack_ID == '4268602' & City == 'National' ~ Units * 1.6 * 1.08, 
           Prod_Desc == 'JANUVIA' & City == '南京' ~ Units * 0.8, 
           Prod_Desc == 'JANUMET' & City == 'National' ~ Units * 3, 
           Prod_Desc == 'DIAMICRON' & City %in% c('杭州', '南京', '上海', '苏州') ~ Units * 2, 
           Prod_Desc == 'DIAMICRON' & City == 'National' ~ Units * 1.1, 
           Prod_Desc == 'VICTOZA' & City == 'National' ~ Units * 2.5, 
           TRUE ~ Units
         ), 
         DosageUnits = case_when(
           Pack_ID == '4268604' & City == '福州' ~ DosageUnits * 1.8, 
           Pack_ID == '4268604' & City == '广州' ~ DosageUnits * 0.7, 
           Pack_ID == '4268602' & City == '广州' ~ DosageUnits * 4, 
           Pack_ID == '4268604' & City == '上海' ~ DosageUnits * 1700 * 0.8, 
           Pack_ID == '4268602' & City == '上海' ~ DosageUnits * 4100 * 0.8, 
           Pack_ID == '4268602' & City == '苏州' ~ DosageUnits * 1.2, 
           Pack_ID == '4268604' & City == 'National' ~ DosageUnits * 1.08, 
           Pack_ID == '4268602' & City == 'National' ~ DosageUnits * 1.6 * 1.08, 
           Prod_Desc == 'JANUVIA' & City == '南京' ~ DosageUnits * 0.8, 
           Prod_Desc == 'JANUMET' & City == 'National' ~ DosageUnits * 3, 
           Prod_Desc == 'DIAMICRON' & City %in% c('杭州', '南京', '上海', '苏州') ~ DosageUnits * 2, 
           Prod_Desc == 'DIAMICRON' & City == 'National' ~ DosageUnits * 1.1, 
           Prod_Desc == 'VICTOZA' & City == 'National' ~ DosageUnits * 2.5, 
           TRUE ~ DosageUnits
         )) %>% 
  # left_join(adj.factor, by = c("Prod_Desc", "Pack_ID", "City")) %>% 
  # mutate(factor = if_else(is.na(factor), 1, factor),
  #        Sales = round(Sales * factor, 2),
  #        Units = round(Units * factor),
  #        DosageUnits = round(DosageUnits * factor)) %>% 
  # select(-factor) %>% 
  arrange(Date, Province, City, Pack_ID)

write.xlsx(msd.result, "03_Outputs/06_MSD_CHC_OAD_2020Q2.xlsx")


##---- Dashboard ----
# MSD history
msd.history <- read.xlsx('02_Inputs/MSD_Dashboard_Data_2019_2020Q1_20200721.xlsx', 
                         sheet = 2)

# DPP4
add.dpp4 <- msd.history %>% 
  filter(Prod_Desc %in% c('JANUVIA', 'ONGLYZA', 'FORXIGA', 'GALVUS', 'TRAJENTA', 
                          'VICTOZA', 'NESINA', 'JANUMET', 'EUCREAS'), 
         Date == '2020Q1') %>% 
  mutate(Date = '2020Q2', 
         Value = case_when(
           City %in% c('Beijing', 'Shanghai') & Prod_Desc == 'TRAJENTA' ~ Value * 1.185, 
           City %in% c('Beijing', 'Hangzhou', 'Nanjing', 'Shanghai') & Prod_Desc == 'VICTOZA' ~ Value * 1.247, 
           City == 'Fuzhou' & Prod_Desc == 'NESINA' ~ Value * 1.089, 
           City %in% c('Hangzhou', 'Shanghai') & Prod_Desc == 'ONGLYZA' ~ Value * 1.045, 
           City == 'Shanghai' & Prod_Desc == 'JANUMET' ~ Value * 3, 
           City == 'Shanghai' & Prod_Desc == 'FORXIGA' ~ Value * 1.924, 
           TRUE ~ NaN
         )) %>% 
  filter(!is.na(Value))

# dashoboard info
dly.dosage <- read.xlsx("02_Inputs/OAD_PDot转换关系.xlsx", startRow = 4)
dly.dosage1 <- read.xlsx("02_Inputs/PROD_NAME_Not_Matching1.xlsx")
dly.dosage2 <- read.xlsx("02_Inputs/PROD_NAME_Not_Matching2.xlsx")
msd.category <- read.xlsx("02_Inputs/MSD交付匹配.xlsx", cols = 1:2)
city.en <- read.xlsx("02_Inputs/MSD交付匹配.xlsx", cols = 4:7)

dly.dosage3 <- dly.dosage %>% 
  mutate(PROD_NAME = gsub(" \\S*$", "", PROD_NAME)) %>% 
  distinct()

dly.dosage.pack <- msd.result %>% 
  mutate(PROD_NAME = paste0(Prod_Desc, " ", Pck_Desc), 
         PROD_NAME = gsub("\\s+", " ", str_trim(PROD_NAME))) %>% 
  left_join(dly.dosage3, by = "PROD_NAME") %>% 
  bind_rows(dly.dosage1, dly.dosage2) %>% 
  filter(!is.na(DLY_DOSAGE)) %>% 
  distinct(Pack_ID, DLY_DOSAGE)

# MSD dashboard
msd.dashboard <- msd.result %>% 
  mutate(PROD_NAME = paste0(Prod_Desc, " ", Pck_Desc), 
         PROD_NAME = gsub("\\s+", " ", str_trim(PROD_NAME))) %>% 
  left_join(dly.dosage.pack, by = "Pack_ID") %>% 
  left_join(msd.category, by = "ATC3") %>% 
  left_join(city.en, by = c("Province", "City")) %>% 
  mutate(Province = Province_EN, 
         City = City_EN, 
         PDot = DosageUnits / DLY_DOSAGE) %>% 
  setDT() %>% 
  melt(id.vars = c("Channel", "MKT", "Date", "Province", "City", "ATC3", 
                   "Category", "Molecule_Desc", "Prod_Desc", "Pck_Desc", 
                   "Pack_ID", "Corp_Desc"),
       measure.vars = c("Sales", "Units", "DosageUnits", "PDot"),
       variable.name = "Measurement",
       value.name = "Value",
       variable.factor = FALSE) %>% 
  bind_rows(msd.history[msd.history$Date == '2019', ], add.dpp4) %>% 
  mutate(Value = round(Value)) %>% 
  arrange(Channel, Date, Province, City, MKT, Pack_ID)

write.xlsx(msd.dashboard, '03_Outputs/06_MSD_CHC_OAD_Dashboard_2020Q2.xlsx')

# MSD dashboard
msd.dashboard.m <- msd.result %>% 
  mutate(PROD_NAME = paste0(Prod_Desc, " ", Pck_Desc), 
         PROD_NAME = gsub("\\s+", " ", str_trim(PROD_NAME))) %>% 
  left_join(dly.dosage.pack, by = "Pack_ID") %>% 
  left_join(msd.category, by = "ATC3") %>% 
  left_join(city.en, by = c("Province", "City")) %>% 
  mutate(Province = Province_EN, 
         City = City_EN, 
         PDot = DosageUnits / DLY_DOSAGE) %>% 
  setDT() %>% 
  melt(id.vars = c("Channel", "MKT", "Date", "Province", "City", "ATC3", 
                   "Category", "Molecule_Desc", "Prod_Desc", "Pck_Desc", 
                   "Pack_ID", "Corp_Desc"),
       measure.vars = c("Sales", "Units", "DosageUnits", "PDot"),
       variable.name = "Measurement",
       value.name = "Value",
       variable.factor = FALSE) %>% 
  filter(Date == '2020Q2') %>% 
  bind_rows(msd.history, add.dpp4) %>% 
  mutate(Value = round(Value)) %>% 
  arrange(Channel, Date, Province, City, MKT, Pack_ID)

write.xlsx(msd.dashboard.m, '03_Outputs/06_MSD_CHC_OAD_Dashboard_2020Q2_m.xlsx')


