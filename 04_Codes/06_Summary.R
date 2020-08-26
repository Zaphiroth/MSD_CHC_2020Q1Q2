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

# MSD history
msd.history <- read.xlsx('02_Inputs/MSD_Dashboard_Data_2019_2020Q1_20200717.xlsx')

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
         Corp_Desc = if_else(Corp_Desc == "LVYE GROUP", "LUYE GROUP", Corp_Desc)) %>% 
  group_by(Pack_ID, Channel, Province, City, Date, ATC3, MKT, Molecule_Desc, 
           Prod_Desc, Pck_Desc, Corp_Desc) %>% 
  summarise(Sales = sum(Sales, na.rm = TRUE),
            Units = sum(Units, na.rm = TRUE),
            DosageUnits = sum(DosageUnits, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(Sales = round(Sales, 2),
         Units = round(Units),
         DosageUnits = round(DosageUnits)) %>% 
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
  # left_join(adj.factor, by = c("Prod_Desc", "Pack_ID", "City")) %>% 
  # mutate(factor = if_else(is.na(factor), 1, factor),
  #        Sales = round(Sales * factor, 2),
  #        Units = round(Units * factor),
  #        DosageUnits = round(DosageUnits * factor)) %>% 
  # select(-factor) %>% 
  arrange(Date, Province, City, Pack_ID)

write.xlsx(msd.result, "03_Outputs/06_MSD_CHC_OAD_2020Q2.xlsx")


##---- Dashboard ----
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
  bind_rows(msd.history[msd.history$Date == '2019', ])

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
  bind_rows(msd.history)

write.xlsx(msd.dashboard.m, '03_Outputs/06_MSD_CHC_OAD_Dashboard_2020Q2_m.xlsx')


