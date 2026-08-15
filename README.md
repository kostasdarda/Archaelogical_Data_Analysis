# Neanderthal Hunting: Understanding Spatiotemporal Variation in Zooarchaeological Assemblages

This repository contains the dataset, processing workflows, and analytical code for evaluating spatial and temporal variations in Neanderthal (*Homo neanderthalensis*) foraging and prey choice strategies across Eurasia (200 ka to 20 ka BP). 

This work was conducted as part of the **STAMP** project at the Globe Institute (University of Copenhagen) and the **NeanderEdge** project at Aarhus University.

---

## 📌 Project Overview & Objectives

Understanding Paleolithic hunting strategies often involves testing whether foragers acted as megafauna specialists or flexible opportunistic hunters. This project evaluates Neanderthal prey choice by applying the **Prey Choice Model (PCM)** from Behavioral Ecology to a large zooarchaeological dataset.

### Key Objectives:
1. **Database Construction & Standardization:** Process and standardize taxonomic data from Paleolithic sites across Eurasia dating between 200 ka and 20 ka BP.
2. **Prey Ranking:** Classify target genera into profitability rank categories (**High**, **Medium**, **Low/Megafauna**) based on Post-Encounter Return Rates (PERR).
3. **Temporal Analysis:** Evaluate shifts in prey proportions across archaeological periods (Middle vs. Upper Paleolithic) and binned temporal sequences using LOESS smoothing and Generalized Additive Models (GAMs).
4. **Statistical Testing:** Test PCM predictions regarding diet expansion using Generalized Linear Models (GLMs).
5. **Ecological Integration:** Pair zooarchaeological prey abundances with species-specific habitat suitability models derived from Ecological Niche Models (ENMs), specifically for reindeer (*Rangifer tarandus*).

---

## Dataset Summary

The primary dataset is derived from the **Role of Culture in Early Expansions of Humans Out of Africa Database (ROAD)**.

- **Initial Dataset:** 9,422 zooarchaeological assemblages across Western Eurasia.
- **Filtering Criteria:** Restricting observations to contexts with absolute dates (radiometric, luminescence, amino acid, ESR) and valid NISP counts for class *Mammalia*.
- **Taxonomic Standardization:** Focused on **21 target genera**. Proportional abundances were calculated by excluding rodent remains (*Rodentia*) to eliminate background noise:
  $$\text{Standardized Proportion} = \frac{\text{Genus NISP}}{\text{Total Mammalia NISP} - \text{Rodentia NISP}}$$


### Prey Rank Classification (12 Target Taxa):

| Rank Category | Genera | Notes / Description |
| :--- | :--- | :--- |
| **High** | *Rangifer*, *Sus*, *Dama*, *Cervus* | Highest post-encounter return rates (most profitable) |
| **Medium** | *Alces*, *Equus*, *Bison*, *Saiga* | Intermediate returns |
| **Low / Megafauna** | *Mammuthus*, *Coelodonta*, *Megaloceros*, *Palaeoloxodon* | Lower returns or significantly higher handling costs |

---
