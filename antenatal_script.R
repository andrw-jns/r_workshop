library(tidyverse)
# For creating tables in PowerPoint:
library(flextable)
library(officer)


# Exercise 1: ----------------------------------------------------------------
# Do  rural hubs (Bridgnorth, Robert Jones) serve older mothers? 


# Boxplot of ages Outpatient activity - antenatal scans by site.

scans <- read_csv("antenatal_scans.csv")

scans %>% 
  filter(flag_shropshire == 1) %>% 
  ggplot(aes(site_name, age))+
  geom_boxplot()+
  # slightly easier than rotating labels:
  coord_flip()


# Exercise 2: ----------------------------------------------------------------

# How many Shropshire registered patients go to sites outside Shropshire?
# (We'd like to see flows inside shropshire as well)

# Create dataframes of patient flows to the sites (numbers and percentages) and
# turn the dataframes to flextables and then send these to PowerPoint

# Take a quick look:
scans %>% 
  group_by(site_name) %>% 
  summarise(n = n()) %>% 
  arrange(desc(n))

# Inside shropshire:
tab1 <- scans %>% 
  group_by(site_name, flag_shropshire) %>% 
  summarise(n = n()) %>%
  # ungroup: otherwise the mutate - below - will use the above groups
  ungroup %>% 
  arrange(desc(n)) %>% 
  # Add a columnn for percentages
  mutate(perc = n/sum(n)*100) %>% 
  filter(flag_shropshire == 1) %>% 
  # We do not wish the flag in our final table:
  select(- flag_shropshire)

# Outside shropshire
tab2 <- scans %>% 
  group_by(site_name, flag_shropshire) %>% 
  summarise(n = n()) %>%
  # ungroup: otherwise the mutate - below - will use the above groups
  ungroup %>% 
  arrange(desc(n)) %>% 
  mutate(perc = n/sum(n)*100) %>% 
  filter(flag_shropshire == 0) %>% 
  # We do not wish the flag in our final table:
  select(- flag_shropshire)

  # 

# Demonstration:
# To PowerPoint via Flextables (proof of concept) -------------------------
# We'll need to contruct 'flextables' than package Officer can work with
# Send above dataframes to the function flextable, and add some basic formatting

flex1 <- tab1 %>% 
  flextable() %>% 
  fontsize(size = 8) %>%
  bold(part = "header") %>% 
  fontsize(size = 8, part = "header") %>%
  theme_alafoli() %>%
  autofit() 


flex2 <- tab2 %>% 
  flextable() %>% 
  fontsize(size = 8) %>%
  bold(part = "header") %>% 
  fontsize(size = 8, part = "header") %>%
  theme_alafoli() %>%
  autofit() 


# 4. PowerPoint -----------------------------------------------------------

# Create a powerpoint object (default styles)
pres <- read_pptx()

# layout_properties ( x =pres, layout = "Two Content", master = "Office Theme")

pres %>% 
  add_slide(layout = "Two Content", master = "Office Theme") %>%
  # To set title text:
  ph_with_text(type = 'title', str = "Patient flows for antenatal scans") %>% 
  # ph_with_flextable(ft_b1, type = "body", index = 1) %>%
  ph_with_flextable(flex1, type = "body", index = 1) %>%
  ph_with_flextable(flex2, type = "body", index = 2) %>%
  # ph_with_flextable(ft_b1, type = "body", index = 3) %>% 
  print(target = "antenatal.pptx")
