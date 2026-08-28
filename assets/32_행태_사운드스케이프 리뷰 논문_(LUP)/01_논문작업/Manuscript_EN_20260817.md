# Soundscape and human behaviour in urban open space: a systematic review and meta-analysis of a bidirectional relationship

**Target journal**: Landscape and Urban Planning
**Registration**: OSF Registries, osf.io/7ew8q (Generalized Systematic Review Registration)
**Draft**: 2026-08-17 (English version) · corpus 98 studies

---

## Author information

Hyun In Jo¹\*

¹ Department of Architectural Engineering, Hanyang University, Seoul 04763, Korea

**[Corresponding author]**

\*Send correspondence to: Hyun In Jo (best2012@hanyang.ac.kr)

Department of Architectural Engineering

Hanyang University

222 Wangsimni-ro, Seongdong-gu

Seoul 04763, Republic of Korea

**CRediT authorship contribution statement**

Hyun In Jo: Conceptualization, Methodology, Software, Formal analysis, Data curation, Writing, original draft, Writing, review & editing, Visualization.

**Declaration of competing interest**

The authors declare no competing financial interests or personal relationships that could have appeared to influence the work reported in this paper.

**Funding**

This research received no specific grant from any funding agency in the public, commercial, or not-for-profit sectors.

**Ethics**

This study analysed only publicly available, aggregated and de-identified data and did not involve human participants directly; institutional review board approval was therefore not required.

**Data availability**

All review-level materials are openly available. The registered protocol is at osf.io/7ew8q. Extraction tables, verdict tables, item-level MMAT appraisals, meta-analytic input tables, sensitivity and supplementary analyses, the protocol deviation log, and complete lists of included and full-text-excluded reports are provided as Supplementary S1–S14. Full texts of included papers are under publisher copyright and are not redistributed; each is identified by its DOI.

---

## Abstract

Soundscape research in urban open space has accumulated for over two decades around perceptual appraisal, yet its link to the outcome that planning actually manages, **observable behaviour**, has never been synthesised systematically. Under a pre-registered protocol (osf.io/7ew8q) we systematically reviewed and meta-analysed this relationship, treating as eligible both the forward pathway (sound changes behaviour) and the reverse pathway (activity shapes the acoustic environment). Searches of Web of Science, Scopus and PubMed (2 August 2026), citation tracking and a supplementary index yielded 98 included studies. Quality was appraised with MMAT 2018; effect sizes were pooled with Hartung–Knapp-adjusted REML random-effects models. All four behavioural clusters pointed in the predicted direction, but only two mean effects excluded zero, social interaction (*k* = 4, *g* = +0.68, 95% CI +0.20 to +1.16) and the perception–behaviour correlation (*k* = 6, *r* = +0.43, 95% CI +0.19 to +0.61), and **the 95% prediction interval included zero in all four clusters.** The widely cited walking-speed effect rested entirely on low-quality studies, most of which manipulated headphone audio rather than the sound of the space. Of 187 directional records, 40% ran in reverse, and the two directions were nearly balanced for space use, activity and social behaviour. Three generations of behavioural measurement (self-report, systematic observation, sensing) accumulated rather than replaced one another. None of the ten planning levers derived from the corpus reached high confidence. The best-supported action is to control operational noise where interaction is intended; the field currently supplies hypotheses to verify through monitoring, not design standards.

**Keywords**: soundscape; behaviour; urban open space; systematic review; meta-analysis; landscape planning

---

## 1. Introduction

### 1.1 Sound as a planning variable

Landscape architecture and urban planning have long worked with what can be seen. Site plans and bird's-eye views, visual openness and enclosure, planting form and seasonal change make up the core of the design language. The observation that urban design leans towards the visible while sound goes largely unaddressed was made half a century ago [1]. Yet when a person actually sits in a square, what makes the place bearable, or makes them hurry away, is not only what enters the field of view. The lorry passing on the adjacent road, the low-frequency hum of an underground car-park vent, the music played by the shop across the street, and whether the voice of a companion can be heard: these define the character of the place. Sound cannot be avoided by looking elsewhere, cannot be fully screened by a wall, and usually surfaces as a problem only after the design is finished.

The claim that the quality of public space should be judged by what people do is not new. Whyte [2] filmed small plazas in New York and recorded where people sat and where they stopped; Gehl [3] divided outdoor activity into necessary, optional and social activities and observed that optional and social activities grow as the quality of a space rises. In this tradition the unit by which a space succeeds or fails is observed behaviour, not a satisfaction score. But the environmental variables this tradition examined were largely visual and physical; sound was never confronted directly.

This is not to say that planning practice has ignored sound; rather, the way sound has been handled is narrow. Environmental noise regulation is built around a single indicator, the equivalent sound level (LAeq), and zone-specific limit values, and its purpose is to contain health damage and complaints. The evidence behind this framework is ample: the dose–response relationship between traffic-noise exposure and annoyance has been quantified in large-scale syntheses [4], and it is established that noise disturbs sleep and raises cardiovascular risk [5]. The problem is that a separate set of questions falls outside the framework. A design question such as "what should this square sound like if people are to use it" is not answered by a limit value, because being below the limit and being well used are different things.

The concept of the **soundscape** starts precisely from this gap. As proposed by Schafer [6], it treats sound not as a physical quantity to be removed but as the whole sonic environment as people perceive and experience it in context. ISO 12913-1 [7] standardised the concept as the acoustic environment "as perceived or experienced and/or understood by a person or people, in context", settling the language of the field.

### 1.2 What the field has established, and what it has not

What has accumulated in the two decades since standardisation is mainly knowledge about **perception**. Axelsson et al. [8] showed with principal component analysis that soundscape appraisal reduces to two axes, pleasantness and eventfulness, which became the descriptive language of the ISO 12913 series. Aletta et al. [9] organised the components of models that predict these appraisals, and Kang et al. [10] set the field's agenda with ten questions for soundscape research in the built environment. Recently, the association between landscape elements and soundscape perception has even been synthesised meta-analytically [11].

This accumulation carries a marked skew: **the outcome variable is overwhelmingly a rating scale.** Pleasantness, harmony, appropriateness, restorativeness, annoyance and preference are all values that people mark on questionnaires.

Planning decisions are not made about those values. When deciding whether to build a square, where to place benches, which façade should carry the ventilation plant, or in which hours to permit outdoor performance, what is actually at stake is **what people do**. Is the space used? Do people stay, or pass through? Do they speak to strangers? The assumption that ratings and behaviour are connected is widely shared, but studies measuring the connection itself are far fewer, and they have never been synthesised systematically.

This is the gap the present review addresses. Syntheses of perception already exist in numbers; a synthesis of behaviour does not. And the two cannot substitute for each other. A poorly rated sound does not necessarily drive people out of a space, and a sound rated no worse than average can still make conversation impossible and undo the social use of a place. The strength of the perception–behaviour link is a value to be measured, not assumed.

### 1.3 Why now: a measurement threshold has been crossed

There is a reason this synthesis has become possible now: **71% of the eligible studies in this review appeared in 2020 or later** (70 of 98). Five years ago this review could not have been written.

What changed is that the cost of observing behaviour fell sharply. Recording behaviour in outdoor space used to mean observers standing in the field entering head counts into a grid, or recruiting participants for surveys. Samples were small, observation windows short, and comparisons across seasons and weekdays practically impossible. Today GPS trajectories, smartphone location data, crowdsourced exercise logs, video analytics and wearable sensors do the same work at scales several orders of magnitude larger. This corpus includes studies covering 81,403 bicycle passages and 13,322 street segments.

Measuring the acoustic environment became cheaper at the same time. Low-cost Internet-of-Things (IoT) noise-sensor networks, machine-learning models that estimate the acoustic environment from street-view imagery, and smartphone-based soundwalks have spread, making it feasible to record sound and behaviour in the same place at the same time. The surge of literature in the past five years is the product of both measurement costs falling together.

The surge has a price. Much of the new work is observational big data, and observation yields correlation, not causation. As will be seen, **only 19% of this corpus manipulated the acoustic environment in the field.** Data have multiplied; manipulation has not kept pace. That is the current state of the field.

### 1.4 The pathway that runs the other way

Reading the literature, we found a structure we had not expected. The dominant research frame casts sound as **exposure** and behaviour as **outcome**: how does noise move people? Yet of the 187 directional records extracted from this corpus, **74, that is 40%, run the other way.**

Reverse-direction studies ask questions such as these. How much does sound pressure rise as crowd density grows? How does a dance gathering in a square change the acoustic environment of the adjacent housing? Do people alone and people with companions notice the same sounds differently? How should an open-air market be laid out to reduce the noise it generates? These studies treat the soundscape as the outcome and human activity as the cause.

More important, the reverse pathway is not evenly distributed. Behaviours that are easy to manipulate experimentally, such as movement and staying, are dominated by forward designs. For **space use, activity and social behaviour the two directions are nearly balanced.** The more collective and sustained the behaviour, the more researchers have measured how it produces the acoustic environment.

This structure speaks directly to planning practice. What designers control is not only sources of noise: deciding which activities a space will hold is itself deciding what the space will sound like. Where benches are clustered, on which side the playground sits, on which days outdoor performance is permitted; these are acoustic design. A frame that treats sound only as exposure therefore places roughly half of what this literature has measured out of view from the start. This is why the present review sets out from a bidirectional framing.

### 1.5 Positioning against adjacent syntheses

Adjacent syntheses exist, but none covers the scope of this review.

The closest is the review by Wang et al. [12], a systematic review of how urban soundscapes affect physical activity. It stands in the same direction as ours in treating sound and behaviour together, but the single behaviour it covers makes it a subset of the present review; movement, staying, space use and social interaction lie outside its scope.

Zhang et al. [11] meta-analysed the association between landscape elements and soundscape perception. That is the link immediately upstream of ours: it quantifies the connection from landscape to perception, whereas this review addresses the next link, from perception and exposure to behaviour. Placed end to end, the two syntheses connect the chain from landscape through perception to behaviour quantitatively for the first time.

Buxton et al. [13] synthesised the effects of natural sounds, but on national-park data, with a focus on health benefits, and with psychological outcomes such as stress reduction and improved affect. Observed behaviour in urban open space is not covered.

The present review also adds an axis absent from prior syntheses: **how behaviour itself has been measured.** We coded three generations of behavioural measurement (self-report, systematic observation, and sensing and trajectories) and tracked them over time. This is not methodological bookkeeping; it bears directly on interpretation, because an association measured only by self-report cannot be read with the same weight as one measured on observed behaviour, and the high heterogeneity of our correlational cluster very likely contains this factor.

### 1.6 Research questions

Against this background we posed four questions.

**RQ1.** What is the size and direction of the association between the acoustic environment and observable behaviour in urban open space, and does it differ by behavioural domain?

**RQ2.** How much of the evidence runs in reverse (from behaviour to soundscape), and in which domains?

**RQ3.** How has behavioural measurement changed, and does new instrumentation replace older methods or accumulate on them?

**RQ4.** What can landscape and urban planning actually prescribe on this evidence, and with what confidence?

The first three questions describe the literature; the last translates that description into the language of practice. In answering it we held to the principle that the strength of a prescription must not exceed the strength of its evidence. As will be seen, the result is modest.

---

## 2. Methods

### 2.1 Registration and reporting

This review was pre-registered with OSF Registries (osf.io/7ew8q) on the Generalized Systematic Review Registration template before screening began, and reporting follows the PRISMA 2020 statement [14]. Every departure from the registered protocol is recorded in Supplementary S1 and summarised in §2.7, **including those that changed results.** Analysis rules (the effect-size metric, contrast frames, the handling of multiple effects within one paper, pseudoreplication) were fixed in a separate document **before any effect size was computed** (Supplementary S2). The reason for keeping that order is simple: rules set after seeing results follow the results.

### 2.2 Eligibility criteria

Eligibility required all five of the following criteria.

The subjects are humans. Animal behaviour research was excluded, including soundscape ecology; that field shares our terminology while studying entirely different subjects, entered the search results in bulk, and became the single largest exclusion reason in the database branch.

The setting is outdoor or semi-outdoor urban and landscape public space: parks, streets, squares, waterfronts, outdoor campuses, residential open space and recreation areas. Laboratory and virtual-reality studies were eligible only where the stimulus reproduced an outdoor scene; these entered the main analysis, with a parallel sensitivity analysis excluding them. Indoor, in-vehicle, hospital and workplace settings were excluded.

The exposure is the acoustic environment: noise, natural sound, music, or the soundscape as a composite. For the reasons given in §1.4, **the reverse pathway, in which behaviour and activity shape the acoustic environment, was also defined as eligible.** This decision was made at registration; it was not introduced afterwards to enlarge the corpus.

The outcome must be **observable behaviour**: movement (walking speed, route, crossing), staying (dwell time, sitting), space use (visitation, occupancy), social interaction, and physical activity and leisure, with self-reported behaviour also accepted. Studies reporting only perception, preference, annoyance, restorativeness or physiological response were excluded. Studies whose only behavioural outcome is a stated *intention*, for example a single item asking "would you revisit this park", were not excluded but set aside for sensitivity analysis.

Designs were restricted to empirical, peer-reviewed journal articles published in English. This restriction, together with the exclusion of the Chinese-language literature, was fixed at registration; because it bears directly on interpretation it is taken up again in Limitations.

### 2.3 Information sources and search

Web of Science Core Collection, Scopus and PubMed were searched on 2 August 2026. The search string combined three concept blocks with AND: acoustic environment, behaviour, and urban outdoor setting. Full strings are in Supplementary S3. The strategy was validated before execution against five benchmark papers already known to be eligible; all five were retrieved.

Two supplementary identification routes were registered, and both were executed. The first is **citation tracking**: with every included study as a seed, backward references and forward citations were collected from the OpenAlex API. The second treats OpenAlex as a **fourth index**, identifying records absent from the three databases; 352 unique records were obtained and screened.

### 2.4 Selection process

The full flow of the three routes is shown in Fig. 1.

<<FIG:Fig1_PRISMA>>

**Database branch.** After deduplication, 1,316 records were screened at title and abstract against seven numbered exclusion criteria, combining AI pre-classification with human verification. One decision at this stage shaped the results: **records that remained uncertain were carried to full-text assessment rather than excluded.** In this literature a behavioural outcome is often measured yet never mentioned in the abstract; papers whose abstracts report only perceptual outcomes while the behaviour appears in a table in the body are in fact numerous.

**Citation-tracking branch.** The backward references and forward citations of the 84 studies included at that point yielded 2,073 records not already in the screening pool. Citation tracking, unlike the main search, has very low topical specificity: the reference lists of included papers carry large volumes of statistics, greenspace-and-health, and general planning literature. The registered three-block search logic was therefore applied to titles as a priority filter, deprioritising 1,645 records that matched at most one block or belonged to animal acoustics. The remaining 428 were screened by title, then 146 by abstract, and full texts were sought. Because this automated prioritisation was not specified at registration, it is logged as a departure (S1, D1-3).

**Supplementary-index branch.** Of the 352 unique OpenAlex records, 23 duplicated the two pools above; 329 were screened and 2 proved eligible. The yield is small, but this branch left the review's clearest warning about index metadata: of the five full texts retrieved, **two were recorded as `language=en` yet were written in Korean and Japanese**, with only their titles and abstracts in English. Document type and language cannot be adjudicated from index metadata alone; had we trusted the metadata, we would have violated our own registered language criterion.

**Boundary rules.** Three types of boundary case that recurred at full text were fixed as rules before individual verdicts and then applied uniformly. Studies pairing residential noise exposure with non-specific physical activity were excluded; studies whose only outcome is behavioural intention were set aside for sensitivity analysis; and reverse-direction studies with rating outcomes were included, because the registered definition makes the reverse pathway eligible. A fourth rule, assigning laboratory reproductions of outdoor scenes to sensitivity-only, was introduced during the citation-tracking branch and applied only there; the cross-branch inconsistency is logged (S1, D6-1), and §3.3 shows that it has no quantitative consequence.

### 2.5 Data collection and quality appraisal

A single reviewer extracted data from full texts into sixteen fixed fields: setting, design, sample, exposure, behavioural domain, measurement method, direction, key findings, and **verbatim effect statistics with their location in the source.** Effect statistics were transcribed exactly as printed, and **no conversion or estimation was performed at extraction.** The separation is not a formality: as §2.7 reports, when extraction and pooling pass through the same hands this boundary can blur, and in our case it did.

Methodological quality was appraised with the **Mixed Methods Appraisal Tool (MMAT) 2018** [15]. MMAT assigns each study to one of five categories (qualitative, randomised, non-randomised quantitative, quantitative descriptive, mixed methods) and asks five questions per category, each judged met (Yes), not met (No) or Can't tell. **Can't tell means that the paper does not contain the information needed to judge, not that the study is poor**; as will be seen, most deficits in this corpus are of this type.

The registration planned RoB 2 in parallel for randomised studies, but the two randomised studies in the corpus report neither their randomisation procedure nor blinding, so applying a second instrument would have produced no new information. It was not applied, and the departure is logged (S1, D2-2).

**One caveat about the appraisal deserves emphasis.** Full texts were machine-extracted, and an initial extraction cap truncated 13 of the 16 citation-tracked studies. MMAT scores unreported items as Can't tell; **truncation is therefore indistinguishable from non-reporting and biases scores downward.** Every truncated study was re-appraised on its complete text: at least one item changed in seven studies, tiers moved in four, and in three of those by two tiers. Before re-appraisal, the citation-tracked literature appeared distinctly weaker than the database yield; the difference was an artefact of the extraction cap, not a property of the literature. We report this because any review using automated full-text extraction is exposed to the same artefact, which to our knowledge has not been reported (S12).

### 2.6 Synthesis

Pooling required at least three independent estimates of the same behavioural outcome under the same contrast frame.

Results reported on different scales need a common unit. Differences between two conditions were converted to **Hedges' *g*** [16], which expresses the difference between two groups relative to the spread within them; by convention 0.2 reads as small, 0.5 as medium and 0.8 as large. Because *g* inflates in small samples, the small-sample correction was applied. Studies reporting 2 × 2 counts of helpers and non-helpers entered through the odds ratio and the Chinn transformation (*d* = ln(OR) × √3/π) [17]. Correlations were pooled on Fisher's *z* and back-transformed to *r*.

The pooling model is **REML random effects with the Hartung–Knapp adjustment**. A random-effects model assumes that the true effect differs somewhat from study to study, which fits a literature this varied in settings and stimuli. The Hartung–Knapp adjustment guards against confidence intervals that come out too narrow when studies are few [18]. Between-study variance is reported as τ², and the share of total variation due to between-study variance, heterogeneity, as *I*²; a large *I*² signals that the studies are measuring different things.

Where one study contributed more than one estimate to the same cluster and contrast frame, one was chosen by a fixed priority: the most direct and objective measurement first, then the primary analysis, then the full sample. Remaining equal-rank effects were averaged within the paper assuming a between-effect correlation of ρ = 0.5. Where two papers share a sample, only one entered any pool. In pseudoreplicated designs the participant *n*, not the observation *n*, was used.

Because between-study variance was large in two clusters, we report the **95% prediction interval** alongside the confidence interval [19]. The two intervals answer different questions, and the distinction decides much of this review's conclusion.

The confidence interval says where the mean of the studies collected so far lies. The prediction interval says what the next single study is expected to find. The former can exclude zero while the latter includes it, and in this review that is exactly what happened. Planning decisions rest on the latter interval: what a designer wants to know is not whether the average so far differed from zero, but what will happen if the prescription is applied.

**Reporting bias could not be assessed.** The registered rule restricted funnel plots and Egger tests to clusters with *k* ≥ 10, and the largest cluster here has *k* = 6. The threshold follows standard guidance, which discourages asymmetry tests on few studies because their power is low [20]. No test was therefore run in any cluster. Small-study effects **can neither be detected nor excluded.**

**Subgroup analysis was registered but not performed.** Splitting clusters of *k* = 3–6 by setting, design or measurement generation leaves one to four estimates per subgroup, below the pre-specified minimum of three, and at *k* = 2 the Hartung–Knapp interval becomes effectively unbounded. Sources of heterogeneity were examined narratively instead.

Eight sensitivity analyses were run: leave-one-out; excluding MMAT low studies; recomputing on observation *n*; excluding rank correlations; excluding inputs that required reconstruction; an alternative definition of the walking-speed pool; excluding laboratory reproductions of outdoor scenes; and, added post hoc, excluding the citation-tracking branch. The last axis asks whether the conclusions depend on the supplementary search route; as will be seen, they do. Full results are in S8.

### 2.7 Departures from the registered protocol

**28** departures from the registered protocol are recorded in full in Supplementary S1 and summarised in Table A below. **Ten changed a reported result.** We list all of them rather than a selection, because a deviation log that reports only comfortable deviations is not a deviation log. Six of the ten need explanation.

**First, MMAT categories were reassigned from full-text reading** (D2-1). Categories pre-assigned from extracted design strings often disagreed with the full text, and because the category determines which items apply, tiers changed. The most common misassignment treated "objective acoustic measurement plus survey" as mixed methods; both components are quantitative, so the quantitative category applies.

**Second, the music-stimulus study was excluded from the walking-speed pool** (D3-6). Including it moves the estimate from *g* = −0.474 (*p* = .326) to −0.742 (*p* = .082); **exclusion is the conservative choice, not the convenient one.** It was excluded because both the exposure (music is added sound) and the behaviour (browsing a shopping street, not transit walking) differ from the rest of the pool.

**Third, we found and corrected a sign-coding error of our own** (D3-4). The social-interaction cluster initially returned *I*² = 93.6%; tracing that implausibly large value showed that the odds ratio of Moser [21] had been coded in the wrong direction. After correction, *I*² = 43.5%. We record this because **the heterogeneity statistic worked as an error detector**, a practical argument for reporting heterogeneity prominently at small *k*.

**Fourth, we excluded a newly retrieved effect that would have strengthened a significant cluster** (D5-4). CT0414 (Nanjing, MMAT high) contrasts an "auditory space" with visual and other sensory spaces rather than sound with sound, and its observation count (2,249) exceeds its units (1,167), so independence fails. Including it moves *p* from .020 to .005 and *I*² from 49% to 77%. It is reported as a sensitivity analysis only.

**Fifth, the extraction cap changed four quality tiers** (D5-3), as described in §2.5.

**Sixth, the results narrative was rewritten when the social-interaction cluster crossed *p* = .05** (D5-6). The first synthesis claimed that *all three condition-contrast clusters include zero and only observational associations accumulate*. Retrieving one 1975 field experiment moved the cluster to *p* = .020, and the claim no longer held. **The claim changed because the result changed, not the reverse**, and the earlier claim is preserved verbatim in S1 for comparison.

The remaining four are of a different kind: errors in our own implementation, exposed when the meta-analysis code was handed to an external verifier, who had not written the manuscript, with instructions to refute it (D7-1 to D7-4). None reverses a conclusion, but two are not the kind of thing to fix quietly.

The heaviest is that **the τ² estimator was maximum likelihood although it was reported as REML.** The function name and the reporting said REML; the equation it solved was the ML score equation. The cause was one function duplicated across three files, since consolidated into a single module. The old estimator understated τ² by 33% (MA1), 80% (MA2), 68% (MA3) and 17% (MA4), and **prediction intervals were accordingly reported too narrow.** All four clusters were recomputed.

Second, **the walking-speed pool violated our own one-effect-per-study rule** (D7-2): the two experiments of one study had entered the same contrast frame separately. Combining them within the paper as pre-specified moved *k* from 4 to 3 and *p* from .179 to .326. Third, **two correlational inputs did not match their sources** (D7-3): one effect had been constructed at the pooling stage although the extraction table recorded `NR`, and no such combination exists in the source; the other used a sample size that was not the analytic unit. The former was removed from the pool; for the latter, the variance was back-calculated from the published confidence interval. Fourth, **the rank-correlation sensitivity had excluded nothing** (D7-4), because the source field was overwritten while reading inputs and then used as the filter; the previous report of "identical to the main analysis" was the product of the bug.

**Table A.** All departures from the registered protocol, generated from the deviation log (Supplementary S1). "Yes" marks a departure that changed a reported result.

| # | Section | Departure | Effect on results |
|---|---|---|---|
| D1-1 | Search and selection | All records still uncertain after title/abstract screening were carried to full text rather than excluded | No — widened retrieval, criteria unchanged |
| D1-2 | Search and selection | Three recurring boundary cases formalised as rules after registration, then applied uniformly | Moved 3 studies to sensitivity-only |
| D1-3 | Search and selection | Title-level priority filter applied to citation-tracking output (registered three-block logic) | Coverage limit; stated in Limitations |
| D1-4 | Search and selection | Citation-tracking branch initially stopped at 'reports sought'; completed by manual retrieval of 71 reports | Resolved — branch now complete (15 studies) |
| D2-1 | Quality appraisal | MMAT categories reassigned from full-text reading rather than from the extracted design string | **Yes** — category determines items, so tiers changed |
| D2-2 | Quality appraisal | RoB 2 not applied separately: only two randomised studies, neither reporting procedure or blinding | No |
| D2-3 | Quality appraisal | One 1978 scanned report left category-undetermined (no text layer) | No — contributes no effect |
| D3-1 | Meta-analysis | Five registered clusters reduced to four; space-use effects were unpoolable (no variance information) | Cluster set; studies remain in narrative synthesis |
| D3-2 | Meta-analysis | Analysis rules (multiple-effect priority, pseudoreplication, non-convertible inputs) fixed in a separate document before any computation | No — pre-specified |
| D3-3 | Meta-analysis | Three unreported values reconstructed (equal-split n; SD back-calculated from p; SE→SD) | Bounded by sensitivity ⑤ |
| D3-4 | Meta-analysis | A sign-coding error of our own found and corrected (I² 93.6% → 43.5%) | **Yes** — corrected before any reporting |
| D3-5 | Meta-analysis | Two sensitivity axes only partially executable (observation-n available for one study; two studies report group-level n only) | Stated per axis |
| D3-6 | Meta-analysis | Walking-speed pool defined without the music-stimulus study | **Yes** — exclusion is the conservative direction (g −0.500 vs −0.742) |
| D4-1 | Reporting | Conceptual framework placed in Discussion rather than Results | No — presentation only |
| D4-2 | Reporting | English-language, journal-only restriction retained as registered | No — registered, not a deviation; consequential for interpretation |
| D5-1 | Citation-tracking branch | Exclusion code for language added (Chinese-language journal indexed with English metadata) | One verdict changed |
| D5-2 | Citation-tracking branch | Intention-outcome rule stated inconsistently across branch instructions | No — no verdict depended on it |
| D5-3 | Citation-tracking branch | Full-text extraction cap truncated 13 of 16 citation-tracked studies; all re-appraised on complete text | **Yes** — 7 items and 4 tiers changed |
| D5-4 | Citation-tracking branch | Four cluster-membership judgements on newly retrieved studies, including one exclusion that would have strengthened the result | **Yes** — see §2.7(4) |
| D5-5 | Citation-tracking branch | Two publication pairs share samples; one member of each pair admitted to any pool | Prevents double-counting |
| D5-6 | Citation-tracking branch | Result narrative revised after the social-interaction cluster crossed p = .05 | **Yes** — claim changed because the result changed |
| D6-1 | Supplementary-index branch | Laboratory-reproduction rule applied to the citation-tracking branch only | No — no such study is poolable |
| D6-2 | Supplementary-index branch | Registered supplementary index (OpenAlex) screened late, after self-audit found it unscreened | No — the 2 studies found contribute no effect |
| D6-3 | Supplementary-index branch | Index language metadata found unreliable (2 of 5 retrieved texts not in English despite `language=en`) | Two verdicts changed |
| D7-1 | Post-hoc statistical verification | τ² estimator was maximum likelihood although reported as REML; corrected and consolidated into a single module | **Yes** — all four clusters recomputed |
| D7-2 | Post-hoc statistical verification | Walking-speed pool violated our own one-effect-per-study rule; the two experiments of one study combined as pre-specified | **Yes** — k 4 → 3, p .179 → .326 |
| D7-3 | Post-hoc statistical verification | Two correlational inputs did not match the source: one effect had been constructed at pooling although extraction recorded none; one used a sample size that was not the analytic unit | **Yes** — k 7 → 6, r .409 → .425 |
| D7-4 | Post-hoc statistical verification | The rank-correlation sensitivity excluded nothing because the source field was overwritten before filtering | **Yes** — real result k = 4, r = .484 |

---

## 3. Results

### 3.1 Study selection

The database search returned 2,073 records (Web of Science 1,010, Scopus 850, PubMed 213). After removing 757 duplicates, 1,316 were screened, 1,127 excluded, and 189 full texts sought. 89 could not be retrieved, 84 of which had been classed as uncertain at screening. 100 full texts were assessed, 16 excluded and 3 reserved for sensitivity analysis, leaving **81 studies** from this route.

Tracking the citations of these studies yielded 2,073 records not already in the pool. After the priority filter, 428 were screened by title and 146 by abstract, and 79 full texts were sought. 54 were retrieved and assessed, 38 excluded and 1 reserved for sensitivity analysis, so this route added **15 studies**, 19% of the database yield. The supplementary index screened 329 records, sought 15, assessed 5 and contributed **2 studies**.

The composition of the supplementary-index exclusions is itself informative. Of 314 exclusions, **147 (47%) were non-empirical**, including book reviews, editorials, commentary and obituaries. This is because the OpenAlex index is far larger than Web of Science and Scopus [22], and it has been reported that using OpenAlex as a search source returns far more records than conventional sources, with a correspondingly larger screening burden [23]. A wider index holds more of the relevant literature, and more of what must be screened out. That two studies were added from outside the three databases shows, at minimum, that **the corpus was not closed by database searching alone.**

The final corpus is **98 studies**, with 4 further studies reserved for sensitivity analysis (102 analysed in total).

The two large routes diverged instructively in yield. Records whose abstracts could be retrieved mechanically, allowing the behavioural outcome to be checked, survived at **46%** (17 of 37); records passed on title alone because no abstract could be retrieved survived **not at all** (0 of 17). **Abstract availability, not topical plausibility, was strongly associated with final inclusion.** That title-only first-stage screening is less sensitive than screening titles together with abstracts has also been reported in methodological work [24]. The two groups were not randomly assigned, however, and the latter numbers only 17, so the influence of the 22 unretrieved title-judged records on the corpus cannot be excluded.

### 3.2 Study characteristics

The corpus is best summarised as **young and skewed**. 70 of the 98 studies (71%) appeared in 2020 or later, the median publication year is 2022, and the oldest study dates from 1975. Across the 56 studies reporting participant counts, the median is 124 (IQR 30–400). The rest report units other than participants (street segments, GPS trajectories, street-view sample points), the largest being 81,403 bicycle passages. That sample size no longer has a single unit is itself a mark of the transition this literature is undergoing.

Settings are dominated by streets (41) and parks (32), with squares (10), campuses, residential open space and waterfronts filling the remainder. Full characteristics are in Table 1.

**Table 1.** Characteristics of the 98 included studies. Percentages are of 98 studies; categories marked † allow a study to appear more than once. The full study-level listing (identifier, route, year, journal, title, country, setting, design, sample, exposure, behavioural domain, measurement generation, direction, quality) is Supplementary S11.

| Characteristic | Studies | % |
|---|---|---|
| **Publication period** | | |
| ≤2009 | 5 | 5% |
| 2010–2019 | 23 | 23% |
| 2020–2026 | 70 | 71% |
| **Study design** | | |
| Mixed observation + survey | 34 | 35% |
| Survey | 25 | 26% |
| Field or natural experiment | 19 | 19% |
| Field observation | 7 | 7% |
| Sensor / big data | 6 | 6% |
| Laboratory / VR experiment | 5 | 5% |
| Other / not reported | 2 | 2% |
| **Setting** † | | |
| Street | 41 | 42% |
| Park | 32 | 33% |
| Square / plaza | 10 | 10% |
| Other | 8 | 8% |
| Campus | 5 | 5% |
| Residential open space | 4 | 4% |
| Waterfront | 4 | 4% |
| **Country** † | | |
| China | 38 | 39% |
| Spain | 7 | 7% |
| United Kingdom | 6 | 6% |
| Germany | 5 | 5% |
| Italy | 4 | 4% |
| Australia | 4 | 4% |
| Czechia | 4 | 4% |
| France | 3 | 3% |
| Other countries (n = 25) | 42 | — |
| **Direction of relationship** | | |
| forward | 57 | 58% |
| reverse | 33 | 34% |
| both | 8 | 8% |
| **Behavioural measurement generation** † | | |
| G1 | 54 | 55% |
| G2 | 42 | 43% |
| G3 | 24 | 24% |
| **Methodological quality (MMAT 2018)** | | |
| High | 22 | 22% |
| Moderate | 40 | 41% |
| Low | 36 | 37% |

G1 = self-report; G2 = systematic observation; G3 = sensing, GPS, video or big-data measurement of behaviour.

Geographic concentration is the corpus's largest constraint. **38 of the 98 studies (39%) were conducted in China**, followed by Spain (7), the United Kingdom (6) and Germany (5). Cross-cultural studies have reported that soundscape appraisal differs across countries even for the same type of urban open space [25,26], and such comparisons are themselves delicate, because results can turn on the translation and cross-cultural adaptation of the perceptual attributes [27]. Either way, pooled estimates must not be read as culturally general values.

<<FIG:Fig7_GeoTime>>

Fig. 2b shows a second, less obvious asymmetry: **31 of the 33 reverse-direction studies appeared in 2016 or later.** The bidirectional evidence base emphasised in §1.4 is younger still than the corpus as a whole, and its accumulation correspondingly shallower.

Most consequential for the conclusions is the design distribution. **Only 19 studies (19%) manipulated the acoustic environment in the field.** The modal design is mixed observation and survey (34), followed by survey (25). This literature observes the acoustic environment far more often than it changes it. Observational research is not the problem in itself; but what planning wants to know is "what changes when something is changed", so the scarcity of manipulation directly constrains prescriptive reach.

### 3.3 Methodological quality

The MMAT tiers of the 98 included studies are high 22, moderate 40 and low 36. That distribution, however, misses what matters most about this corpus. Two much sharper patterns appear at item level (Fig. 3).

<<FIG:Fig8_Quality>>

First, **the two randomised studies receive no credit for their design.** Neither reports how randomisation was performed, whether the groups were equivalent at baseline, or whether assessors were blinded, so all three items are judged Can't tell and the studies meet one and two criteria respectively. The strongest design in the corpus receives among the weakest appraisals purely for lack of reporting. They are not the absolute bottom: four studies meet no criterion at all.

Second, generalising that observation, **the largest deficit is not confirmed deficiency of conduct but deficiency of reporting.** Sample representativeness is clearly met in 10 of 51 quantitative descriptive studies (20%), low non-response bias in 16 of 51 (31%), and adequate control of confounding in 9 of 22 non-randomised studies (41%). These items score low not because the studies were shown to be biased, but because **the information needed to judge is absent.** A field that reports sound pressure to three decimal places routinely omits response rates; the asymmetry is on full display here. The dose–response between noise and annoyance could be quantified at scale because that literature standardised the reporting of exposure and response together [4]; behavioural research has not reached that stage.

One procedural check completes the picture. Excluding the five studies that reproduced outdoor scenes in VR changes no pooled estimate, because none of them contributes a poolable effect; they enter the narrative synthesis only. This also resolves the cross-branch inconsistency noted in §2.4 (D6-1): the rule was applied to one branch only, but no such study enters any pool, so there is no quantitative consequence.

### 3.4 Meta-analyses

Four clusters met the pooling requirement. The characteristics of the contributing studies and their individual effects are in Table 3. The conclusion first: **all four clusters point in the theoretically predicted direction, and none establishes that direction with dependable precision.** Two clusters exclude zero for the mean effect; the prediction interval, which speaks to what the next study will find, includes zero in all four.

**Table 2.** Pooled estimates for the four behavioural clusters. REML random effects with the Hartung–Knapp adjustment. CI = confidence interval (precision of the mean); PI = 95% prediction interval (range expected for a new study). **Every prediction interval includes zero.**

| Cluster | *k* | Metric | Estimate | 95% CI | 95% PI | *p* | *I*² |
|---|---|---|---|---|---|---|---|
| Walking speed (natural sound vs noise) | 3 | *g* | −0.474 | −2.055, +1.108 | −9.13, +8.19 | .326 | 81.4% |
| Staying / dwell time (positive sound vs control) | 3 | *g* | +0.343 | −0.091, +0.778 | −1.67, +2.36 | .077 | 54.4% |
| Social interaction (natural/quiet vs noise) | 4 | *g* | **+0.679** | +0.201, +1.158 | −0.42, +1.78 | **.020** | 48.6% |
| Soundscape perception ↔ behaviour | 6 | *r* | **+0.425** | +0.191, +0.613 | −0.24, +0.82 | **.006** | 92.0% |

**Table 3.** Characteristics of the studies contributing to each meta-analytic cluster. ▲ marks studies retrieved by citation searching. *n* is the analytic sample for the pooled contrast, which can differ from the study's total sample; per-effect values, variances and source locations are Supplementary S7. Effects are Hedges' *g* with 95% CI, except the final cluster, which reports *r*. NR = not reported (variance back-calculated from the published CI).

| No. | Study | Country | Setting | Design | Contrast or measure | *n* | MMAT | Effect [95% CI] |
|---|---|---|---|---|---|---|---|---|
| **Walking speed (Hedges' g; negative = faster under noise)** | | | | | | | | |
| 1 | Franěk 2018 | Czechia | street | field experiment | birdsong vs traffic noise | 57 | low | −1.02 [−1.58, −0.47] |
| 2 | Franěk 2019 | Czechia | street | field experiment | birdsong vs city noise (Exp 1 and 2) | 102 | low | −0.63 [−1.12, −0.13] |
| 3 | Berkouk 2020 | Algeria | street | observational | natural vs traffic sound | 54 | low | +0.23 [−0.31, +0.76] |
| **Staying / dwell time (Hedges' g; positive = longer stay)** | | | | | | | | |
| 4 | Aletta 2016 | United Kingdom | campus | field experiment | music vs no music | 596 | high | +0.39 [+0.19, +0.59] |
| 5 | Ba & Kang 2020 | China | street | field experiment | music vs no sound | 97 | moderate | +0.61 [+0.21, +1.02] |
| 6 | Fu 2026 | China | mixed | mixed | natural sound index (high vs low) | 241 | moderate | +0.21 [+0.07, +0.36] |
| **Social interaction (Hedges' g; positive = more interaction)** | | | | | | | | |
| 7 | Chen 2023 | China | park | field experiment | natural vs noise (group interaction) | 73 | moderate | +0.98 [+0.49, +1.46] |
| 8 | Chen 2024 | China | residential | field experiment | natural vs noise (paired interaction) | 146 | high | +0.55 [+0.22, +0.88] |
| 9 | Moser 1988 | France | street | field experiment | quiet vs roadworks noise (helping) | 150 | high | +0.43 [+0.13, +0.73] |
| 10 | Mathews & Canon 1975 ▲ | NR | laboratory (outdoor scene) | field experiment | quiet vs lawnmower noise (helping) | 80 | moderate | +1.07 [+0.45, +1.69] |
| **Soundscape perception and behaviour (r)** | | | | | | | | |
| 11 | Guo 2024 | China | park | survey | pleasantness with static behaviour | 419 | moderate | +0.56 [+0.49, +0.63] |
| 12 | Zhou 2026 | China | street | mixed | natural sound events with queuing | 315 | low | +0.21 [+0.10, +0.31] |
| 13 | Mansouri 2025 | Algeria | street | mixed | sound comfort with walking comfort | NR | moderate | +0.40 [+0.04, +0.67] |
| 14 | Bao 2023 | China | park | survey | dwell time with restorativeness | 180 | moderate | +0.55 [+0.44, +0.65] |
| 15 | Montes González 2023 ▲ | Spain | street | survey | LAeq with vocal effort | 29 | moderate | +0.65 [+0.37, +0.82] |
| 16 | Cao & Kang 2021 ▲ | United Kingdom | square | survey | companionship with sound noticing | 301 | moderate | +0.16 [+0.05, +0.27] |

<<FIG:Fig2_Forest>>

All four estimates point the way theory predicts (Fig. 4): noise speeds passage, positive sound lengthens stays, quiet and natural sound increase social interaction, and appraisal covaries with behaviour at moderate strength. **How precisely each direction is established differs sharply across clusters, and that difference is the heart of this section.**

Social interaction reached significance on the strength of **two field experiments thirteen years apart**. In Mathews and Canon [28], a lawnmower with its muffler removed raised ambient sound from about 50 to 87 dB(C), and the share of passers-by helping to pick up dropped items fell from 20 of 40 to 5 of 40 (OR = 7.00, *d* = +1.07). Moser [21] applied the same 2 × 2 logic on a Paris street using roadworks noise and reached the same sign. The two studies used unrelated noise sources and different helping paradigms; that they agree carries more information than *k* = 4 alone suggests. Because Mathews and Canon do not state where their data were collected, however, **we do not claim cross-country replication.**

**Even so, the significance is conditional on one study retrieved by citation searching.** The full leave-one-out: removing Chen and Kang [29] gives *g* = +0.581 (*p* = .056); removing Chen et al. [30], +0.775 (*p* = .070); removing Moser [21], +0.803 (*p* = .043); removing Mathews and Canon [28], +0.596 (*p* = .057). **Three of the four removals cross back over *p* = .05.** The cluster is significant as a set rather than through any single study, and the study whose removal returns it to its pre-retrieval state is precisely the one the supplementary search recovered. This is a limitation of the result and, at the same time, evidence that **the registered supplementary search route was not a formality.**

The correlational cluster behaves in the opposite way. It is insensitive to search route (*r* = .425 with citation tracking, .445 without) and to any single study (leave-one-out range *r* = .372 to .484), but its heterogeneity is very high (*I*² = 92.0%), as expected of a pool that mixes forward and reverse pathways and several behavioural outcomes. What it supports is that perception and behaviour broadly move together; which perception moves with which behaviour, and by how much, it cannot say.

**The prediction intervals qualify both significant results.** The two intervals that exclude zero do so for the **mean effect** only; the interval within which a new study's effect is expected to fall **includes zero in all four clusters**: −0.42 to +1.78 for social interaction, *r* = −0.24 to +0.82 for the correlational cluster. Neither result establishes that the next study will carry the same sign. Reporting bias compounds this: with a maximum *k* of 6, below the pre-specified threshold of *k* ≥ 10, neither funnel plots nor Egger tests were run, and small-study effects can neither be detected nor excluded. The walking-speed prediction interval (−9.13 to +8.19), with one degree of freedom at *k* = 3, carries essentially no information; we print it unchanged rather than conceal that this cluster cannot predict anything.

**The walking-speed cluster is the weakest in the corpus, and its weakness is instructive.** All three contributing effects come from MMAT low studies, so excluding low quality leaves nothing to pool (*k* = 0). The corpus holds three high-quality walking studies, and none enters the pool: one manipulated music tempo rather than environmental sound [31], one did not test noise against speed directly [32], and one delivered augmented footstep sounds through headphones [33]. **A claim widely cited in this literature, that noise makes people walk faster, rests on some of its least stable evidence.**

### 3.5 Direction of the relationship (RQ2)

Of the 212 domain-level records extracted from the 98 studies, 113 are forward (sound → behaviour), 74 reverse (behaviour → sound) and 25 bidirectional. **Reverse records are 40% of directional records**: 74 of the 187 that are forward or reverse, or 35% if bidirectional records enter the denominator. The unit here is the study × behavioural-domain record, not the study.

<<FIG:Fig4_Direction>>

The reverse pathway is not evenly spread, and that unevenness is the point of this section (Fig. 5). Movement (33 forward to 8 reverse) and staying (16 to 8) are dominated by forward designs; these are behaviours that are easy to manipulate: put headphones on people and have them walk, or switch on speakers and time their stays. Space use (26 to 23), activity (18 to 21) and social behaviour (20 to 14) sit near parity.

Where behaviour is collective and sustained, researchers have measured how it *produces* the acoustic environment. Crowd density predicts sound pressure, activity programming brings broadcast music, and companionship changes what is noticed. These studies do not fit the dominant frame that casts sound as exposure, but from the standpoint of planning practice they are the more direct ones, because what designers place includes not only sound sources but **activity** itself. A one-way frame discards nearly half of what this literature has measured.

### 3.6 Measurement generations (RQ3)

The answer to RQ3 is unambiguous: **accumulation, not succession** (Fig. 6).

<<FIG:Fig5_Methods>>

Behavioural measurement by sensors, GPS, video and big data grew from 4 studies in 2010–2019 to 20 from 2020 onwards, a fivefold rise. Over the same interval, self-report grew from 14 to 39 and systematic observation from 11 to 28. **New instrumentation did not push out the older methods; it stacked on top of them.** 22 studies use two or more generations at once.

This matters for two reasons. One is that the field's **basis for triangulation is widening**: studies combining two or more generations rose from 5 in the 2010s to 17 from 2020. Where the movement a GPS trace shows and the movement a participant reports can be compared within one study, things invisible to either alone become visible. The other is a caution. As generations accumulate, **results that measure the same concept in different ways come to coexist in one literature**, and a synthesis that does not separate them leaves much of its heterogeneity unexplained. The *I*² = 92% of our correlational cluster very likely contains this factor.

### 3.7 Coverage and gaps

The evidence map's behavioural domain × sound source combinations are filled except for one source (Fig. 7).

<<FIG:Fig3_EvidenceMap>>

The exception is **aircraft noise**. Three of the five domains (movement, staying, social behaviour) are empty, and the two filled cells hold four domain-level records from two studies in total. This is no accidental gap. Aircraft noise is among the most densely studied topics in environmental noise research: the systematic reviews underpinning the WHO Environmental Noise Guidelines for the European Region treat aircraft noise as an object of synthesis in its own right for annoyance [34] and for sleep [35], and the picture for health effects generally is no different [5]. Yet **within our eligibility criteria, those two studies are all that measure what people do in the outdoor spaces around airports.** That behavioural evidence is thinnest exactly where noise exposure is most severe shows the frame within which this field has been confined.

---

## 4. Discussion

### 4.1 A bidirectional framework for soundscape and behaviour

The conceptual framework this review arrives at is presented in Fig. 8.

<<FIG:Fig6_Framework>>

Two arcs close one loop. The **forward** arc runs acoustic environment → appraisal → behaviour and is what the field has presupposed. The **reverse** arc runs activity and occupancy → sound production → acoustic environment and accounts for 40% of directional records. Between the two sit the moderators that decide whether a given sound produces approach or avoidance: setting type, visual–acoustic congruence, purpose of stay and cultural context, factors that soundscape-perception research has confirmed repeatedly [10,9].

What we would emphasise in this framework is that behaviour is not treated as a single outcome. Behaviour is an **engagement gradient**: avoiding, passing, lingering, interacting, appropriating. The same space is a thoroughfare to one person and a destination to another, and what planning tries to change is usually a position on this gradient. Placing the pooled estimates on the gradient makes the shape of the field visible at a glance: **the evidence is weak at both ends, and the strongest quantitative signal sits at 'interacting'.**

The framework's value is **diagnostic** rather than descriptive. It shows that this literature produces nearly half of its measurements on the reverse arc while organising itself around the forward arc, and that the behaviour gradient has been sampled with little regard to where planning actually intervenes. 'Appropriating', for instance, is close to the ultimate aim of public-space design, and it is among the cells with the thinnest quantitative evidence in this corpus.

### 4.2 What the evidence supports, and how strongly (RQ4)

We translated the corpus into ten **planning levers**. A lever is something a designer or operator can actually move. "The noise is loud" is a description of state, not a lever; "move the ventilation fan to the opposite façade" and "shift mowing to the morning" are levers. The distinction exists so that the review's results can be carried into practice documents as they stand.

Each lever carries a confidence grade, set by five things considered together: how many studies contribute; their MMAT composition; whether a pooled interval excludes zero; whether the effect is observed across different settings; and whether the result survives sensitivity analysis.

**Table 4.** Ten planning levers derived from the corpus, with confidence grading. Confidence combines the number of contributing studies, their MMAT composition, whether a pooled interval excludes zero, diversity of settings, and behaviour under sensitivity analysis. **No lever reaches high confidence.** `Mixed` under Direction means that studies disagree in sign, not that the lever has several effects. Full evidence, caveats and study lists are Supplementary S14.

| Lever | What is changed | Behavioural outcome | Direction | Studies | MMAT mix | Confidence |
|---|---|---|---|---|---|---|
| **L1** Programmed music in public space | Sound draws attention, pulls people towards the source and slows them, extending stay | Staying (dwell time); space use (approach to source); movement (slower wandering) | promotes | 7 | high 3 · mod 2 · low 2 | **moderate** |
| **L2** Natural sound provision (water and birdsong) | Natural sound raises perceived restoration and safety, encouraging talk and lingering | Social interaction; staying (dwell time); vitality of activity | promotes | 6 | high 2 · mod 3 · low 1 | **moderate** |
| **L3** Quiet routes for walking and cycling | Noise avoidance shifts route and mode choice, moving trips onto quieter alignments | Movement (route choice, mode choice, cycling volume) | promotes | 8 | high 4 · mod 1 · low 3 | **moderate** |
| **L4** Remove or reschedule mechanical plant noise | Unpleasant machinery noise creates avoidance routes and acceleration, cutting stay and talk | Social interaction; staying; space use (removal of avoidance routes) | promotes | 7 | high 2 · mod 3 · low 2 | **moderate** |
| **L5** Quiet side and designated quiet zones | A quiet façade or designated zone lowers the psychological barrier to outdoor stay and walking | Activity (walking, exercise, rest); space use; staying | mixed | 10 | high 2 · mod 4 · low 4 | **low** |
| **L6** Acoustic zoning and enclosure of functions | Separating and enclosing noise-generating functions changes the density and interaction of adjacent activity | Space use (crowd density); social interaction (frequency, duration); staying | mixed | 6 | high 1 · mod 5 | **low** |
| **L7** Programming sound-generating public activities | Human sound and activity attract watching and joining, converting passage into stay | Staying (watching, lingering); social interaction; space use | promotes | 6 | high 1 · mod 2 · low 3 | **low** |
| **L8** Auditory guidance and warning signals | Directional signal sound directly adjusts crossing trajectory and response timing | Movement (crossing accuracy, detection timing, smoothness of deceleration) | mixed | 4 | high 1 · mod 2 · low 1 | **low** |
| **L9** Slowing pedestrian pace by natural sound | Noise provokes avoidance and speeds passage; natural sound is assumed to reverse it | Movement (walking speed) | mixed | 5 | high 1 · low 4 | **very low** |
| **L10** Speech-interference criteria for siting social space | Noise interferes with speech, forcing conversation to stop or voices to rise, and eventually deterring verbal interaction altogether | Social interaction (conversation duration, vocal effort); acceptance of verbal interaction | promotes | 3 | high 2 · mod 1 | **low** |

**How to read the grades.** *moderate* (4) — Include in a design proposal, with post-occupancy monitoring · *low* (5) — Test as a hypothesis; do not write into a standard or guideline · *very low* (1) — No prescriptive basis at present.

**No lever reaches high confidence.** Four are moderate (programmed music, natural sound provision, quiet routes, control of mechanical plant noise), five are low, and one is very low. This distribution is itself the review's central finding for practice: **the field currently supports hypotheses to be verified through monitoring, not design standards.**

Two attribution judgements matter more than the grades. Both block readings that are plausible but unsupported by our data.

**First, the social-interaction result is attributable neither to adding natural sound nor to removing noise.** Reading the significant cluster as a case for water features and birdsong is convenient; so is reading it as a case for silencing machinery. Neither is supported. Split the four effects by exposure and the means are nearly identical (natural +0.76, mechanical +0.75); remove the roadworks study and the pooled estimate does not fall but **rises** (+0.68 → +0.80). More fundamentally, the two natural-sound studies contrast birdsong and water against traffic and construction noise: they add pleasant sound and remove noise **at the same time**. Because **no study in the corpus contrasts natural sound against quiet**, the two mechanisms cannot be identified separately from anything in it. What can be said is narrower: the increment that pushed this cluster over the significance threshold came from one mechanical-noise experiment recovered by citation searching.

**Second, the walking-speed literature is narrower than it appears from outside.** Beyond the quality collapse already reported, two structural features emerged during appraisal. One is the mode of delivery: most of the pooled effects manipulated audio delivered over **headphones** during walking, not the sound of the space, and headphones differ fundamentally from environmental sound in masking, personal choice and the allocation of attention. The other is concentration of provenance: the walking evidence is effectively **one research programme**. Two of the three pooled effects come from the same Czech group [36,37]; the two newly added high-quality walking papers are by the same group on the same 1.75–1.8 km circuit, two of them sharing the same belt-camera annotation protocol. The remaining effect is an Algerian field observation with the opposite sign [38].

One study outside the pool changed the sound of a real street. Ba et al. [39] crossed traffic noise at 55.6 and 70.5 dB(A) with plant-scent concentration under covert observation. Mean crowd speed was about 1.14 m/s in the quieter condition and 1.21 m/s in the noisier, **about 0.07 m/s faster under noise**, the direction the literature predicts; the authors report a maximum difference of 0.06 m/s between measurement grids. The study is MMAT low. In sum, **the only evidence from a manipulated real environment yields an effect of a few centimetres per second.** A claim replicated mostly through headphones, in one city on one circuit, is not a replicated finding, and the *k* of a meta-analysis does not reveal this.

### 4.3 Implications for landscape planning

The practical reading of Table 4, in order of defensibility.

**Start with operational noise.** The best-supported behavioural consequence in this corpus is that mechanical noise suppresses social behaviour, including help offered to strangers. The prescription is practically attractive not only because the evidence is comparatively strong. Ventilation plant, mowing schedules and construction hours sit **within routine operational control**; changing them costs little, requires no redesign and needs no budget approval. That two independent field experiments with different noise sources support it can be added.

**Treat added sound as a staying tool, not a draw.** Music lengthens the stay of people already present. The only high-quality study that tested attraction [40] found no effect on the number of people stopping. Programming justified by growth in visitor numbers outruns what the evidence shows. A common error of practice lies here: when a square-activation project sets "more users" as its performance indicator and deploys an acoustic intervention, indicator and evidence part company.

**Introduce a conversability criterion for siting social space** (lever L10, new in this revision). Where people are meant to talk (bench clusters, terraces, meeting places), the operative threshold is not annoyance but whether **conversation holds**, that is, speech interference. The evidence supports a continuous slope, not a cut-off: roughly 0.25 points of self-reported vocal effort per decibel [41]. That is enough to compare candidate sites and not enough to set a standard. Expect about one point of difference on a speech-interference scale between two sites whose equivalent sound levels differ by 4 dB; that is all this evidence can say.

**Do not use walking speed as a design performance indicator**, and do not carry headphone results into public space, for the reasons of the previous section.

The prescriptions we judge **unsupported** are listed in the notes to Table 4. Two are common enough in practice to single out. Masking traffic noise with water features to increase use: **not one study in this corpus tests it on behaviour.** The evidence that natural sound improves perceptual and psychological indicators is real [13], but improved perception and increased use are different propositions, and this prescription has not yet crossed from perception research into behaviour. And treating noise reduction as a route to physical activity is not merely unsupported but **contradicted**: the two largest datasets here report positive associations between noise and activity [42,43], because activity concentrates where cities are loud. Quiet places do not generate activity; active places generate noise.

### 4.4 Methodological implications

**This field measures sound far better than it measures people.** The largest quality deficit is not confirmed design flaws but missing reporting. Sample representativeness is clear in about one study in five, low non-response bias in roughly one in three. The bottom-tier scores of the two randomised studies have exactly the same cause, and it is a fixable one: no new equipment, no new budget, a few lines of reporting.

**Measurement generations accumulate rather than replace.** As §3.6 showed, sensing grew fastest, but self-report and systematic observation grew alongside it, and 22 studies now combine two or more. If studies report enough for the layers to be compared, this is an opportunity.

**Three warnings concern evidence synthesis itself.** All three come from our own work.

First, automated full-text extraction interacts badly with appraisal instruments that score unreported items as Can't tell. Truncated extraction lowered the quality tier of 4 of 16 studies by at least one tier, three of them by two, and **made an entire search route appear to retrieve weaker literature.** The bias runs one way only and is invisible without a deliberate audit. Reviews using machine extraction should report their extraction limits and confirm that appraised studies were appraised on complete text.

Second, index metadata cannot be trusted. Two of the five full texts retrieved through the supplementary index were indexed as English yet written in Korean and Japanese. A review that enforces its language criterion on metadata rather than on documents will violate its own protocol silently.

Third, and most painfully, **computational code cannot be verified by the person who wrote it.** Handing the meta-analysis implementation to an external verifier with instructions to refute it exposed four errors (§2.7). Two are especially telling. **The same τ² estimation function was duplicated across three files, and all three solved maximum likelihood despite their name.** Without the duplication, one fix would have sufficed; with it, three outputs were wrong together. And **a study whose extraction table recorded `NR` for effect statistics had acquired a value at the pooling stage.** When extraction and pooling pass through one person, that boundary blurs. Neither error fell to self-checking; both fell to a verifier who was not the author. Our experience is that **reviews reporting meta-analyses should have their computation re-run independently.**

### 4.5 Research agenda

Of the gaps this corpus leaves, we present six with the largest return on effort, in that order.

**One: a single well-reported walking experiment moves an entire cell.** Nothing special is required. Manipulate the sound of a real space rather than headphones, randomise properly, and report procedure, baseline equivalence and blinding. The current pooled estimate stands on studies that each fail two MMAT criteria, and the capacity to run high-quality walking studies already exists in this field; the three studies seen in §3.4 are the proof. This is a matter of design choice, not capability.

**Two: aircraft noise and behaviour is effectively empty.** Two studies. The health effects of aircraft noise have a vast literature [5]; research around airports has stayed inside the health-and-annoyance frame. What is needed is a design that measures use, routes and stays in the outdoor spaces around airports as behaviour, not on an annoyance scale.

**Three: close the chain from design to acoustic outcome to behaviour.** The natural-sound evidence is almost entirely loudspeaker playback. In the three-link chain in which planting actually produces birdsong and that birdsong changes behaviour, no study verifies the first two links by manipulation. This question has direct stakes for landscape practice, because what designers place is trees, not speakers.

**Four: measure persistence.** Nearly every intervention runs from hours to days, the longest for one season. Whether behavioural change survives the removal of an intervention or fades through habituation, nobody knows. Budget decisions hang exactly here: install-once versus operate-forever changes the cost structure of a project.

**Five: build reverse-direction prediction rules.** Even though the two directions are balanced for space use, activity and social behaviour, few studies provide rules that predict what sound a programme of activities will produce. A handful of equations predict sound pressure from density, but their coefficients cannot be transplanted from one setting to another.

**Six: report the basics.** Many otherwise eligible studies specified no per-group *n*, no standard deviation, or no analytic unit, and entered no pool. That so few effects could be pooled in this review owes less to a shortage of studies than to **insufficient reporting.**

### 4.6 Limitations

The corpus is geographically narrow (China 39%) and young (71% from 2020 onwards); pooled estimates must not be read as culturally general. The English-language, journal-only restriction fixed at registration excludes a literature with substantial Chinese-language contributions, and this interacts with the geographic skew: with Chinese studies at 39% of the corpus while Chinese-language work is excluded, the Chinese research we see is the subset that chose to publish in English.

Across the three routes, 122 full texts could not be retrieved (databases 89, citation tracking 25, supplementary index 8). Retrieval reached 92% for citation-tracked records whose abstracts confirmed a behavioural outcome; the unretrieved records were overwhelmingly title-judged, and none of the title-judged records that were retrieved proved eligible (§3.1). This observation does not eliminate the influence of the unretrieved reports, but it indicates its direction.

Screening and extraction were performed by a single reviewer working with AI assistance and stepwise verification, not by two independent reviewers. Two pairs of papers share samples; one member of each pair entered any pool.

The limits of the quantitative synthesis are as stated repeatedly above. Reporting bias could not be assessed in any cluster, *k* never reaching the pre-specified threshold of 10, so small-study effects remain a live possibility. The registered subgroup analyses could not be run for the same reason: clusters of three to six estimates cannot be split. **With prediction intervals including zero in all four clusters, no significant result should be read as a forecast of what a new study will find.** Finally, the significance of the social-interaction cluster depends on one study retrieved by citation searching and would not have been observed from database searching alone.

---

## 5. Conclusions

Sound changes what people do in urban open space. The evidence for that claim is unevenly distributed across behaviours, and **its strongest and weakest parts are not where received wisdom suggests.**

Across 98 studies, all four meta-analytic clusters point in the theoretically predicted direction, but only two mean effects exclude zero and **no prediction interval does.** The best-established behavioural consequence of the acoustic environment is **social**: mechanical noise was associated with reduced interaction between strangers in two independent field experiments using different noise sources. By contrast, the widely cited claim that noise makes people walk faster rests on three effects all appraised as low quality, mostly manipulating headphone audio rather than places, and drawn substantially from one research programme on one circuit.

Nearly half of the evidence runs in reverse: people's activity makes the soundscape as much as the soundscape shapes activity. A planning frame that treats sound only as exposure discards this.

For practice, the defensible actions are narrow but real. **Control operational noise** where people are meant to interact. Use added sound to lengthen **stays**, not to draw visitors. Judge the siting of social space by whether **conversation holds**. And do not prescribe from perception research what has never been tested on behaviour.

For research, the most valuable next study is unglamorous: one field experiment that manipulates the sound of a real place, randomises properly, and reports its methods completely. That single study would move a cell of evidence that half a century of citation has not.

---

## References

[1] Southworth M. The Sonic Environment of Cities. Environment and Behavior. 1969;1:49-70. doi:10.1177/001391656900100104

[2] Whyte WH. The Social Life of Small Urban Spaces. Washington (DC): Conservation Foundation; 1980.

[3] Gehl J. Life Between Buildings: Using Public Space. Washington (DC): Island Press; 2011.

[4] Miedema HM, Oudshoorn CG. Annoyance from transportation noise: relationships with exposure metrics DNL and DENL and their confidence intervals. Environmental Health Perspectives. 2001;109:409-416. doi:10.1289/ehp.01109409

[5] Basner M, Babisch W, Davis A, Brink M, Clark C, Janssen S, et al. Auditory and non-auditory effects of noise on health. The Lancet. 2014;383:1325-1332. doi:10.1016/s0140-6736(13)61613-x

[6] Schafer RM. The Soundscape: Our Sonic Environment and the Tuning of the World. Rochester (VT): Destiny Books; 1994.

[7] International Organization for Standardization. ISO 12913-1:2014 Acoustics — Soundscape — Part 1: Definition and conceptual framework. Geneva: ISO; 2014.

[8] Axelsson Ö, Nilsson ME, Berglund B. A principal components model of soundscape perception. The Journal of the Acoustical Society of America. 2010;128:2836-2846. doi:10.1121/1.3493436

[9] Aletta F, Kang J, Axelsson Ö. Soundscape descriptors and a conceptual framework for developing predictive soundscape models. Landscape and Urban Planning. 2016;149:65-74. doi:10.1016/j.landurbplan.2016.02.001

[10] Kang J, Aletta F, Gjestland TT, Brown LA, Botteldooren D, Schulte-Fortkamp B, et al. Ten questions on the soundscapes of the built environment. Building and Environment. 2016;108:284-294. doi:10.1016/j.buildenv.2016.08.011

[11] Zhang R, Ma H, Wang C, Zhang Y, Kang J. The associations between landscape elements and soundscape perception: A meta-analysis. Landscape and Urban Planning. 2025;263:105463. doi:10.1016/j.landurbplan.2025.105463

[12] Wang Y, Wu Y, Qin T, Van de Weghe N, Huang H. Assessing the impact of urban soundscapes on physical activity: insights from a systematic review. Cities & Health. 2026:1-26. doi:10.1080/23748834.2026.2683259

[13] Buxton RT, Pearson AL, Allou C, Fristrup K, Wittemyer G. A synthesis of health benefits of natural sounds and their distribution in national parks. Proceedings of the National Academy of Sciences. 2021;118. doi:10.1073/pnas.2013097118

[14] Page MJ, McKenzie JE, Bossuyt PM, Boutron I, Hoffmann TC, Mulrow CD, et al. The PRISMA 2020 statement: an updated guideline for reporting systematic reviews. BMJ. 2021:n71. doi:10.1136/bmj.n71

[15] Hong QN, Fàbregues S, Bartlett G, Boardman F, Cargo M, Dagenais P, et al. The Mixed Methods Appraisal Tool (MMAT) version 2018 for information professionals and researchers. Education for Information. 2018;34:285-291. doi:10.3233/efi-180221

[16] Hedges LV. Distribution Theory for Glass's Estimator of Effect size and Related Estimators. Journal of Educational Statistics. 1981;6:107-128. doi:10.3102/10769986006002107

[17] Chinn S. A simple method for converting an odds ratio to effect size for use in meta-analysis. Statistics in Medicine. 2000;19:3127-3131. doi:10.1002/1097-0258(20001130)19:22<3127::aid-sim784>3.0.co;2-m

[18] Hartung J, Knapp G. On tests of the overall treatment effect in meta‐analysis with normally distributed responses. Statistics in Medicine. 2001;20:1771-1782. doi:10.1002/sim.791

[19] Higgins JPT, Thompson SG, Spiegelhalter DJ. A Re-Evaluation of Random-Effects Meta-Analysis. Journal of the Royal Statistical Society Series A: Statistics in Society. 2009;172:137-159. doi:10.1111/j.1467-985x.2008.00552.x

[20] Sterne JAC, Sutton AJ, Ioannidis JPA, Terrin N, Jones DR, Lau J, et al. Recommendations for examining and interpreting funnel plot asymmetry in meta-analyses of randomised controlled trials. BMJ. 2011;343:d4002. doi:10.1136/bmj.d4002

[21] Moser G. Urban stress and helping behavior: Effects of environmental overload and noise on behavior. Journal of Environmental Psychology. 1988;8:287-298. doi:10.1016/s0272-4944(88)80035-5

[22] Culbert JH, Hobert A, Jahn N, Haupka N, Schmidt M, Donner P, et al. Reference coverage analysis of OpenAlex compared to Web of Science and Scopus. Scientometrics. 2025;130:2475-2492. doi:10.1007/s11192-025-05293-3

[23] Stansfield C, Dehdarirad H, Thomas J, Mathew S, O'Mara‐Eves A. Analyzing the Utility of OpenAlex to Identify Studies for Systematic Reviews: Methods and a Case Study. Cochrane Evidence Synthesis and Methods. 2025;3. doi:10.1002/cesm.70038

[24] Teo L, Van Elswyk ME, Lau CS, Shanahan CJ. Title-plus-abstract versus title-only first-level screening approach: a case study using a systematic review of dietary patterns and sarcopenia risk to compare screening performance. Systematic Reviews. 2023;12. doi:10.1186/s13643-023-02374-3

[25] Deng L, Kang J, Zhao W, Jambrošić K. Cross-National Comparison of Soundscape in Urban Public Open Spaces between China and Croatia. Applied Sciences. 2020;10:960. doi:10.3390/app10030960

[26] Nguyen TL, Puyoo-Hialle M, Nguyen TTHN. Cultural influences on urban soundscape perception: A comparison of French, Japanese, and Vietnamese participants. Applied Acoustics. 2026;254:111414. doi:10.1016/j.apacoust.2026.111414

[27] Papadakis NM, Aletta F, Kang J, Oberman T, Mitchell A, Stavroulakis GE. Translation and cross-cultural adaptation methodology for soundscape attributes – A study with independent translation groups from English to Greek. Applied Acoustics. 2022;200:109031. doi:10.1016/j.apacoust.2022.109031

[28] Mathews KE, Canon LK. Environmental noise level as a determinant of helping behavior. Journal of Personality and Social Psychology. 1975;32:571-577. doi:10.1037/0022-3514.32.4.571

[29] Chen X, Kang J. Natural sounds can encourage social interactions in urban parks. Landscape and Urban Planning. 2023;239:104870. doi:10.1016/j.landurbplan.2023.104870

[30] Chen X, Kang J, Wang M. The impact of the community's sound environment on social interactions among residents. Building and Environment. 2024;266:112094. doi:10.1016/j.buildenv.2024.112094

[31] Franěk M, van Noorden L, Režný L. Tempo and walking speed with music in the urban context. Frontiers in Psychology. 2014;5. doi:10.3389/fpsyg.2014.01361

[32] Franěk M, Režný L. Environmental Features Influence Walking Speed: The Effect of Urban Greenery. Land. 2021;10:459. doi:10.3390/land10050459

[33] Schrapel M, Happe J, Rohs M. EnvironZen: Immersive Soundscapes via Augmented Footstep Sounds in Urban Areas. i-com. 2022;21:219-237. doi:10.1515/icom-2022-0020

[34] Guski R, Schreckenberg D, Schuemer R. WHO Environmental Noise Guidelines for the European Region: A Systematic Review on Environmental Noise and Annoyance. International Journal of Environmental Research and Public Health. 2017;14:1539. doi:10.3390/ijerph14121539

[35] Basner M, McGuire S. WHO Environmental Noise Guidelines for the European Region: A Systematic Review on Environmental Noise and Effects on Sleep. International Journal of Environmental Research and Public Health. 2018;15:519. doi:10.3390/ijerph15030519

[36] Franěk M, Režný L, Šefara D, Cabal J. Effect of Traffic Noise and Relaxations Sounds on Pedestrian Walking Speed. International Journal of Environmental Research and Public Health. 2018;15:752. doi:10.3390/ijerph15040752

[37] Franěk M, Režný L, Šefara D, Cabal J. Effect of birdsongs and traffic noise on pedestrian walking speed during different seasons. PeerJ. 2019;7:e7711. doi:10.7717/peerj.7711

[38] Berkouk D, Bouzir TAK, Maffei L, Masullo M. Examining the Associations between Oases Soundscape Components and Walking Speed: Correlation or Causation?. Sustainability. 2020;12:4619. doi:10.3390/su12114619

[39] Ba M, Li Z, Kang J. Research on the Combined Effects of Plant Odor and Traffic Noise on Crowd Behaviors in Urban Environments. Landscape Architecture Frontiers. 2024;12:47. doi:10.15302/j-laf-1-020106

[40] Aletta F, Lepore F, Kostara-Konstantinou E, Kang J, Astolfi A. An Experimental Study on the Influence of Soundscapes on People’s Behaviour in an Open Public Space. Applied Sciences. 2016;6:276. doi:10.3390/app6100276

[41] Montes González D, Barrigón Morillas JM, Rey-Gozalo G. Effects of noise on pedestrians in urban environments where road traffic is the main source of sound. Science of The Total Environment. 2023;857:159406. doi:10.1016/j.scitotenv.2022.159406

[42] Dzhambov AM, Burov A, Markevych I, Kostadinov KR, Dimitrova D, Helbich M, et al. Environmental characteristics and physical activity: A cross-sectional study in Bulgaria's five largest cities. International Journal of Hygiene and Environmental Health. 2026;275:114816. doi:10.1016/j.ijheh.2026.114816

[43] Huang D, Tian M, Yuan L. Sustainable design of running friendly streets: Environmental exposures predict runnability by Volunteered Geographic Information and multilevel model approaches. Sustainable Cities and Society. 2023;89:104336. doi:10.1016/j.scs.2022.104336
---

## Supplementary material

| ID | Title | Source file |
|---|---|---|
| **S1** | Protocol deviation log (28 departures, with effect on results) | `deviation_log.md` |
| **S2** | Pre-specified analysis rules (fixed before any effect was computed) | `analysis_rules.md` |
| **S3** | Full search strings for the three databases and the supplementary index | `search_strings_20260802_205110.md` |
| **S4** | PRISMA 2020 flow with all three identification routes; PRISMA checklist | `prisma_flow.md` |
| **S5** | All 98 included studies — identifiers, route, year, journal, title | `references_all_included.md` |
| **S6** | Reports excluded at full text, with reason codes | `ft_verdicts_v2.csv`, `ct_screen_final.csv` |
| **S7** | Meta-analytic input tables (per-effect values, variance, source location) | `ma/ma_*_input.csv` |
| **S8** | Sensitivity analyses (eight axes, full leave-one-out) | `ma/ma_sensitivity_v2.md` |
| **S9** | Prediction intervals, reporting-bias assessment, subgroup feasibility | `ma/ma_supplementary.md` |
| **S10** | MMAT 2018 item-level appraisals for all 102 analysed studies | `quality_detail_v2.csv` |
| **S11** | Study characteristics table (Table 1, extended) | `table1_v2.csv` |
| **S12** | Truncation audit — effect of extraction limits on quality appraisal | `quality_truncation_effect.md` |
| **S13** | Evidence map counts (domain × source, direction × domain, generation × period) | `evidence_map_v2.md` |
| **S14** | Planning-lever matrix with confidence grading (Table 4, extended) | `design_implications_v2.md` |
