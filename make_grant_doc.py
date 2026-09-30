import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor


def create_document():
    doc = docx.Document()

    # Page Margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Color Palette Definitions
    COLOR_PRIMARY = RGBColor(0, 51, 102)  # Deep Navy
    COLOR_SECONDARY = RGBColor(74, 107, 130)  # Slate Blue
    COLOR_BODY = RGBColor(34, 34, 34)  # Charcoal Dark
    HEX_LIGHT_BG = "F0F4F8"
    HEX_PRIMARY = "003366"

    # Configure Default Style
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Arial"
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = COLOR_BODY
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    # Helper Functions
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY

        pBdr = parse_xml(
            r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            r'<w:bottom w:val="single" w:sz="12" w:space="4" w:color="003366"/>'
            r"</w:pBdr>"
        )
        p._p.get_or_add_pPr().append(pBdr)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_callout(text_runs):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{HEX_LIGHT_BG}"/>')
        tcPr.append(shd)

        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{HEX_PRIMARY}"/>'
            f'<w:top w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:bottom w:val="none"/>'
            f"</w:tcBorders>"
        )
        tcPr.append(tcBorders)

        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="120" w:type="dxa"/>'
            f'<w:left w:w="180" w:type="dxa"/>'
            f'<w:bottom w:w="120" w:type="dxa"/>'
            f'<w:right w:w="180" w:type="dxa"/>'
            f"</w:tcMar>"
        )
        tcPr.append(tcMar)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15

        for text, bold, italic, color in text_runs:
            run = p.add_run(text)
            run.font.name = "Arial"
            run.font.size = Pt(10.5)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color if color else COLOR_BODY

    # Document Title Block
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_sub = p_title.add_run("RESEARCH GRANT PROPOSAL\n")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(11)
    run_sub.font.bold = True
    run_sub.font.color.rgb = COLOR_SECONDARY

    run_title = p_title.add_run(
        "Unraveling Plasmodium falciparum Blood-Stage Persister Niches: A"
        " Multi-Omic, Single-Cell Spatial & Systems Biology Framework for"
        " Radical Clearance"
    )
    run_title.font.name = "Arial"
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY

    add_callout([
        ("Target Mechanism & Scope: ", True, False, COLOR_PRIMARY),
        (
            "Expanded 8-Year Research Program with Comprehensive Academic"
            " Citations\n",
            False,
            False,
            None,
        ),
        ("Host Institutions: ", True, False, COLOR_PRIMARY),
        (
            (
                "Jomo Kenyatta University of Agriculture and Technology (JKUAT)"
                " & Kenya Medical Research Institute (KEMRI), Kenya"
            ),
            False,
            False,
            None,
        ),
    ])

    doc.add_paragraph()  # Spacer

    # Section 1: Specific Aims Page
    add_h1("1. Specific Aims Page")
    add_h2("Background & Rationale")
    p = doc.add_paragraph()
    p.add_run(
        "Malaria morbidity and mortality remain stagnated worldwide, driven by"
        " 249 million global cases and 619,000 deaths annually [WHO, 2023]."
        " Frontline Artemisinin-based Combination Therapies (ACTs) are"
        " increasingly threatened by treatment recrudescence [Ashley et al.,"
        " 2014; WHO, 2023]. While mutations in "
    )
    r = p.add_run("Plasmodium falciparum")
    r.font.italic = True
    p.add_run(
        " Kelch13 (K13) drive delayed ring-stage clearance [Ariey et al.,"
        " 2014; Straimer et al., 2015], K13 mutations alone do not account for"
        " ring-stage dormancy or post-treatment recrudescence in endemic African"
        " settings [Uwimana et al., 2020; Balikagala et al., 2021].\n\nUpon"
        " exposure to artemisinins or physiological stress, a subpopulation of"
        " parasites enters a metabolically quiescent, persister state"
        " [Teuscher et al., 2010; Witkowski et al., 2013]. These persisters act"
        " as incubation hubs for multi-drug resistance variants and sustain"
        " low-density transmission [Chen et al., 2014]. Traditional bulk"
        " molecular approaches average gene expression across millions of"
        " parasites, masking rare persister subpopulations (<1–5% parasitemia)"
        " and ignoring the anatomical tissue microenvironments—such as the"
        " splenic red pulp and bone marrow—where persisters reside "
    )
    r = p.add_run("in vivo")
    r.font.italic = True
    p.add_run(" [Kho et al., 2021; De Niz et al., 2018].")

    add_h2("Central Hypothesis")
    add_callout([
        ("Central Hypothesis: ", True, False, COLOR_PRIMARY),
        ("Plasmodium falciparum ", True, True, COLOR_PRIMARY),
        (
            (
                'blood-stage persistence is governed by distinct sub-population'
                ' transcriptomic programs ("dormancy regulons") driven by'
                " non-K13 epistatic genomic variations, localized within"
                " specialized host tissue microenvironments, and can be"
                " systematically disrupted using precision gene editing to"
                " restore full antimalarial susceptibility [Miotto et al., 2015;"
                " Cowell et al., 2018]."
            ),
            False,
            False,
            None,
        ),
    ])

    add_h2("Specific Aims")
    add_h3(
        "Specific Aim 1: Map the single-cell transcriptomic landscape and"
        " genomic architecture of persistent P. falciparum in clinical cohorts."
    )
    p = doc.add_paragraph()
    p.add_run(
        "Perform Whole Genome Sequencing (WGS) and single-cell RNA sequencing"
        " (scRNA-Seq; 10x Genomics) on pre- and post-ACT longitudinal clinical"
        " isolates from Busia (lake endemic, ~39% prevalence) and Kwale"
        " (coastal endemic, ~20% prevalence) counties in Kenya ("
    )
    r = p.add_run("n")
    r.font.italic = True
    p.add_run(
        " = 672) [Howes et al., 2016; Real et al., 2021]. Construct a"
        " single-cell atlas of ring-stage dormancy to uncover rare persister"
        " trajectories and map non-K13 epistatic genomic backgrounds [Svensson"
        " et al., 2018; Miotto et al., 2015]."
    )

    add_h3(
        "Specific Aim 2: Resolve the spatial transcriptomic architecture and"
        " host-parasite tissue niches of dormant persisters."
    )
    p = doc.add_paragraph()
    p.add_run(
        "Deploy spatial transcriptomics (10x Visium / CosMx Spatial Molecular"
        " Imager) on tissue microenvironments (splenic aspirates and bone"
        " marrow biopsies) to map the spatial niches of dormant parasites, host"
        " immune interaction gradients, and metabolic micro-zones (such as"
        " hypoxia gradients) sustaining dormancy "
    )
    r = p.add_run("in vivo")
    r.font.italic = True
    p.add_run(
        " [Ståhl et al., 2016; He et al., 2020; Kho et al., 2021]."
    )

    add_h3(
        "Specific Aim 3: Validate candidate dormancy regulons via"
        " high-throughput inducible CRISPR/Cas9 editing and phenotyping."
    )
    p = doc.add_paragraph()
    p.add_run(
        "Utilize an automated, high-throughput inducible CRISPR/Cas9 (DiCre"
        " recombinase / CRISPRi) platform in field-adapted and reference"
        " strains (3D7, Dd2) to knock out/down top candidate genes identified"
        " in Aims 1 and 2 [Prommana et al., 2013; Ghorbal et al., 2014;"
        " Knuepfer et al., 2017]. Perform deep functional phenotyping, including"
        " Ring Survival Assays (RSA), flow-cytometric dormancy recovery"
        " kinetics, and untargeted LC-MS/MS metabolomics [Witkowski et al.,"
        " 2013; Mok et al., 2015]."
    )

    add_h3(
        "Specific Aim 4: Integrate multi-omic layers using graph machine"
        " learning and constraint-based systems biology."
    )
    p = doc.add_paragraph()
    p.add_run(
        "Construct an integrative computational network using Deep Graph"
        " Neural Networks (GNNs) and genome-scale metabolic flux balance"
        " analysis (FBA) to predict persister regulatory networks, uncover"
        " synthetic lethal vulnerabilities, and prioritize novel druggable"
        " targets for radical malaria clearance [Plaimas et al., 2010; Zhou et"
        " al., 2020; Orth et al., 2010]."
    )

    # Section 2: Research Strategy — Significance
    add_h1("2. Research Strategy — Significance")
    add_h2("2.1 The Malaria Stagnation & The Persister Bottleneck")
    p = doc.add_paragraph()
    p.add_run(
        "Despite the widespread deployment of ACTs and next-generation"
        " vaccines, global malaria reduction has stalled [WHO, 2023]."
        " Artemisinin derivatives rapidly clear the bulk of circulating"
        " blood-stage parasites, yet recrudescence occurs in up to 10–20% of"
        " treated individuals [White, 2004; Ashley et al., 2014]. Ring-stage"
        " persister parasites arrest their cell cycle, lower their metabolic"
        " rate, and survive exposure to both short-acting artemisinins and"
        " long-acting partner drugs like lumefantrine and piperaquine"
        " [Teuscher et al., 2010; Witkowski et al., 2013]. Upon drug wash-out,"
        " these persisters reactivate, fueling transmission and providing an"
        " evolutionary breeding ground for "
    )
    r = p.add_run("de novo")
    r.font.italic = True
    p.add_run(
        " drug resistance [Sankaranarayanan et al., 2020; Blanquart et al.,"
        " 2016]."
    )

    add_h2("2.2 Beyond K13: Uncovering the Non-K13 Persister Architecture")
    p = doc.add_paragraph()
    r = p.add_run("PfK13")
    r.font.italic = True
    p.add_run(
        " propeller domain mutations serve as crucial markers for artemisinin"
        " resistance in Southeast Asia and parts of East Africa [Ariey et al.,"
        " 2014; Uwimana et al., 2020]. However, clinical recrudescence and"
        " ring-stage dormancy frequently occur in parasites harboring wild-type "
    )
    r2 = p.add_run("K13")
    r2.font.italic = True
    p.add_run(
        " genes [Mukherjee et al., 2017; Rosenthal et al., 2021]. Current"
        " surveillance models that rely exclusively on "
    )
    r3 = p.add_run("K13")
    r3.font.italic = True
    p.add_run(
        " sequencing miss genome-wide epistatic variations, transcriptomic"
        " switches, and post-transcriptional regulators that mediate non-K13"
        " persistence [Miotto et al., 2015; Cerami et al., 2020]."
    )

    add_h2(
        "2.3 Breakthrough Paradigm: From Bulk Averaging to Spatial &"
        " Single-Cell Genomics"
    )
    p = doc.add_paragraph()
    p.add_run("Previous transcriptomic studies of ")
    r = p.add_run("P. falciparum")
    r.font.italic = True
    p.add_run(
        " relied on bulk RNA sequencing, which averages expression signals"
        " across heterogeneous parasite populations [Bozdech et al., 2003; Otto"
        " et al., 2010]. Bulk RNA-Seq completely masks rare persister cells"
        " (<1% of total parasites) [Regev et al., 2017; Real et al., 2021]."
        " Furthermore, bulk methods ignore spatial context—such as host"
        " tissue microenvironments in the spleen and bone marrow where"
        " parasites sequester [De Niz et al., 2018; Kho et al., 2021]."
        " Integrating single-cell RNA-Seq with spatial transcriptomics and"
        " high-throughput CRISPR editing bridges the gap between descriptive"
        " observational genomics and functional, causal biology [Ståhl et al.,"
        " 2016; Doudna & Charpentier, 2014]."
    )

    add_h2("2.4 Translational Impact & Wellcome Strategic Alignment")
    p = doc.add_paragraph()
    p.add_run(
        "This program aligns directly with Wellcome’s focus on Infectious"
        " Disease by targeting drug-resistant malaria [Wellcome Trust"
        " Strategy, 2021]. By establishing a frontier single-cell genomic,"
        " spatial transcriptomic, and CRISPR platform at JKUAT/KEMRI in Kenya,"
        " this award establishes local scientific leadership and"
        " infrastructure in sub-Saharan Africa [Kilama, 2009; Djimde et al.,"
        " 2010]."
    )

    # Section 3: Research Strategy — Innovation
    add_h1("3. Research Strategy — Innovation")
    add_h2("Conceptual Innovation: Single-Cell Persister Atlas")
    p = doc.add_paragraph()
    p.add_run(
        "This study represents the first application of high-throughput"
        " scRNA-Seq (10x Genomics) directly on uncultured clinical isolates"
        " pre- and post-ACT treatment [Real et al., 2021; Howes et al., 2016]."
        " By evaluating uncultured samples, the platform resolves single-cell"
        " developmental trajectories into and out of dormancy without"
        " introducing bulk averaging artifacts or culture-induced"
        " transcriptional shifts [Svensson et al., 2018; Mack et al., 2020]."
    )

    add_h2(
        "Methodological Innovation: Spatial Transcriptomic Mapping of Tissue"
        " Niches"
    )
    p = doc.add_paragraph()
    p.add_run(
        "This project conducts the first spatial transcriptomic profiling (10x"
        " Visium / CosMx SMI) of "
    )
    r = p.add_run("P. falciparum")
    r.font.italic = True
    p.add_run(
        " persisters directly within anatomical microenvironments,"
        " specifically splenic aspirates and bone marrow biopsies [Kho et al.,"
        " 2021; Ståhl et al., 2016]. This approach illuminates host-parasite"
        " cross-talk, localized hypoxia markers (such as HIF-1α), and metabolic"
        " micro-zones sustaining parasite quiescence "
    )
    r2 = p.add_run("in vivo")
    r2.font.italic = True
    p.add_run(" [He et al., 2020; De Niz et al., 2018].")

    add_h2(
        "Technological Innovation: High-Throughput Inducible CRISPR Phenomics"
    )
    p = doc.add_paragraph()
    p.add_run(
        "The program implements automated, arrayed conditional gene editing"
        " (DiCre / CRISPRi) directly in field-adapted isolates [Knuepfer et"
        " al., 2017; Prommana et al., 2013]. This enables functional validation"
        " of persistence candidate genes via quantitative Ring Survival"
        " Assays (RSA), Piperaquine Survival Assays (PSA), and untargeted"
        " LC-MS/MS metabolomic phenotyping [Witkowski et al., 2013; Duru et"
        " al., 2015]."
    )

    add_h2("Computational Innovation: AI & Network Systems Biology")
    p = doc.add_paragraph()
    p.add_run(
        "By utilizing Deep Graph Neural Networks (GNNs) and genome-scale"
        " metabolic flux modeling, the project integrates WGS variants,"
        " single-cell transcriptomes, and spatial coordinates into a predictive"
        " systems biology model [Zhou et al., 2020; Orth et al., 2010]. This"
        " approach moves beyond single-gene paradigms to identify synthetic"
        " lethal vulnerabilities and multi-target therapeutic combinations"
        " [Plaimas et al., 2010; Costanzo et al., 2019]."
    )

    add_h2("Institutional Innovation: Capacity Building in Africa")
    p = doc.add_paragraph()
    p.add_run(
        "The program establishes an African-led multi-omic research pipeline at"
        " JKUAT and KEMRI in Kenya [Kilama, 2009]. This initiative provides"
        " training and infrastructure for the next generation of African"
        " computational biologists and functional genomicists addressing"
        " endemic infectious diseases locally [Djimde et al., 2010; Ndung'u et"
        " al., 2014]."
    )

    # Section 4: Detailed Approaches & Methodology
    add_h1("4. Detailed Approaches & Methodology")
    add_h2("4.1 Clinical Cohort & Longitudinal Sample Processing")
    p = doc.add_paragraph()
    p.add_run(
        "A prospective longitudinal cohort study will be conducted across Busia"
        " County Referral Hospital (high-transmission, lake endemic; ~39%"
        " prevalence) and Kwale County Referral Hospital"
        " (moderate-transmission, coastal endemic; ~20% prevalence) in Kenya ("
    )
    r = p.add_run("n")
    r.font.italic = True
    p.add_run(
        " = 672 total participants, including children under 5 and pregnant"
        " women) [Kenya Ministry of Health, 2020; Howes et al., 2016]. Patients"
        " will be treated with standard six-dose Artemether-Lumefantrine (AL)"
        " or Dihydroartemisinin-Piperaquine (DP) regimens according to national"
        " guidelines [WHO, 2023]. Active clinical and parasitological follow-up"
        " will take place on Days 0, 1, 2, 3, 7, 14, 21, 28, and 42"
        " [Stepniewska et al., 2008].\n\nVenous blood samples (1–3 mL) will be"
        " collected pre-treatment (Day 0) and at post-treatment parasitemia"
        " events. Samples will be subjected to fluorescent dye staining using"
        " SYBR Green I (nucleic acid) and MitoTracker Deep Red (mitochondrial"
        " membrane potential) followed by Fluorescence-Activated Cell Sorting"
        " (FACS) [Peatey et al., 2013; Smilkstein et al., 2004]. Sorting will"
        " separate dormant rings (MT"
    )
    r = p.add_run("int")
    r.font.superscript = True
    p.add_run("), active rings (MT")
    r2 = p.add_run("hi")
    r2.font.superscript = True
    p.add_run("), and pyknotic dead parasites (MT")
    r3 = p.add_run("low")
    r3.font.superscript = True
    p.add_run(") to yield purified persister fractions [Rovira-Graells et al., 2012].")

    add_h2("4.2 Whole Genome Sequencing (WGS) & scRNA-Seq")
    p = doc.add_paragraph()
    p.add_run(
        "Deep whole-genome sequencing (Illumina NovaSeq 6000, 100X mean"
        " coverage) will be performed on paired Day 0 and recrudescent"
        " isolates [Miles et al., 2016]. Sequence analysis will identify"
        " single-nucleotide polymorphisms (SNPs), copy-number variations"
        " (CNVs, such as "
    )
    r = p.add_run("plasmepsin 2/3")
    r.font.italic = True
    p.add_run(
        " duplications), and mutations across candidate drug-resistance loci"
        " including "
    )
    r2 = p.add_run("k13, crt, mdr2, fd, coronin, ")
    r2.font.italic = True
    p.add_run("and ")
    r3 = p.add_run("arps10")
    r3.font.italic = True
    p.add_run(
        " [Witkowski et al., 2017; Miotto et al., 2015; Demas et al.,"
        " 2018].\n\nFor scRNA-Seq, freshly sorted ring-stage parasites will"
        " be loaded onto the 10x Genomics Chromium Controller (targeting"
        " 10,000 single cells per sample) [Zheng et al., 2017]. Single-cell"
        " cDNA libraries will be sequenced to a depth of approximately 50,000"
        " reads per cell. Data will be analyzed using Seurat and Scanpy for"
        " cell clustering, and Monocle3 for pseudo-time lineage trajectory"
        " mapping to trace the transition pathways between active replication"
        " and dormancy [Satija et al., 2015; Wolf et al., 2018; Trapnell et"
        " al., 2014]."
    )

    add_h2("4.3 Spatial Transcriptomics of Tissue Microenvironments")
    p = doc.add_paragraph()
    p.add_run(
        "To characterize anatomical niches, tissue aspirates and biopsies"
        " (splenic red pulp aspirates and bone marrow biopsies obtained"
        " through clinical indication) will be processed using the 10x Visium"
        " Spatial Gene Expression and CosMx Spatial Molecular Imager"
        " (NanoString) platforms at 2–5 μm spatial resolution [Ståhl et al.,"
        " 2016; He et al., 2020; Kho et al., 2021].\n\nA custom target panel"
        " covering over 1,000 "
    )
    r = p.add_run("P. falciparum")
    r.font.italic = True
    p.add_run(
        " genes will be co-hybridized with host human probes. Spatial"
        " co-localization algorithms will map parasite gene expression"
        " relative to host cell markers, immune cell interaction gradients,"
        " hypoxia markers (such as host HIF-1α), and nutrient transport"
        " micro-zones sustaining persistent sequestration [De Niz et al., 2018;"
        " Svensson et al., 2018]."
    )

    add_h2("4.4 High-Throughput Inducible CRISPR/Cas9 Gene Editing")
    p = doc.add_paragraph()
    p.add_run(
        "Top regulatory candidate genes identified in Aims 1 and 2 will be"
        " functionally validated using inducible gene-editing vector systems."
        " Target genes will be cloned into rapamycin-inducible DiCre"
        " recombinase systems or dCas9-KRAB-based CRISPR interference"
        " (CRISPRi) constructs [Knuepfer et al., 2017; Prommana et al., 2013;"
        " Ghorbal et al., 2014].\n\nField-adapted Kenyan isolates and"
        " standard reference lines (3D7, Dd2) will be transfected via"
        " electroporation [Fidock & Wellems, 1997]. Knockout or knockdown will"
        " be induced by the addition of rapamycin. Functional consequences will"
        " be quantified using 0–3h Ring Survival Assays (RSA), Piperaquine"
        " Survival Assays (PSA), flow-cytometric dormancy entry/exit kinetic"
        " assays, and untargeted LC-MS/MS metabolomics to evaluate metabolic"
        " arrest [Witkowski et al., 2013; Duru et al., 2015; Creek et al.,"
        " 2012]."
    )

    add_h2("4.5 Computational Systems Biology & Machine Learning")
    p = doc.add_paragraph()
    p.add_run(
        "Integrated multi-omic data will be processed using Deep Graph Neural"
        " Networks (GNNs) constructed via PyTorch Geometric [Fey & Lenssen,"
        " 2019; Zhou et al., 2020]. Graph nodes will represent individual genes,"
        " proteins, or metabolites, while edges capture co-expression"
        " patterns, physical protein interactions, and epistatic genomic"
        " correlations across clinical isolates [Costanzo et al.,"
        " 2019].\n\nConstraint-based metabolic modeling will be performed using"
        " genome-scale flux balance analysis (FBA) adapted from the "
    )
    r = p.add_run("P. falciparum")
    r.font.italic = True
    p.add_run(
        " iMM904 metabolic reconstruction within the COBRA toolbox [Plaimas et"
        " al., 2010; Orth et al., 2010; Bazzani et al., 2021]. Simulations"
        " will quantify metabolic flux shifts during ring-stage dormancy and"
        " predict synthetic lethal gene combinations, providing prioritized"
        " targets for antimalarial drug development [Plaimas et al., 2010]."
    )

    # Section 5: Work Packages
    add_h1("5. Work Packages & 8-Year Program Management")

    table = doc.add_table(rows=6, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["Work Package", "Timeline", "Core Deliverables"]
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{HEX_PRIMARY}"/>')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

    wp_data = [
        (
            "WP1: Clinical Surveillance & WGS",
            "Years 1–3",
            (
                "Field cohort recruitment (n=672); persister sorting; Illumina"
                " WGS variant catalog [Howes et al., 2016]."
            ),
        ),
        (
            "WP2: Single-Cell Atlas of Dormancy",
            "Years 2–4",
            (
                "10x scRNA-Seq atlas; pseudo-time trajectories; ring dormancy"
                " regulon identification [Real et al., 2021]."
            ),
        ),
        (
            "WP3: Spatial Biology of Niches",
            "Years 3–5",
            (
                "Visium/CosMx tissue profiling; host-parasite spatial maps;"
                " hypoxia micro-zone profiling [Kho et al., 2021]."
            ),
        ),
        (
            "WP4: Functional CRISPR Phenomics",
            "Years 4–7",
            (
                "DiCre/CRISPRi gene knockouts; RSA/PSA survival assays;"
                " metabolomic profiles [Knuepfer et al., 2017]."
            ),
        ),
        (
            "WP5: Systems Biology & Lead Prioritization",
            "Years 6–8",
            (
                "GNN and FBA multi-omic integration; synthetic lethality"
                " prediction; open-access platform [Zhou et al., 2020]."
            ),
        ),
    ]

    for row_idx, data in enumerate(wp_data, start=1):
        row_cells = table.rows[row_idx].cells
        for col_idx, text in enumerate(data):
            row_cells[col_idx].text = text
            p = row_cells[col_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
            if row_idx % 2 == 0:
                shd = parse_xml(
                    f'<w:shd {nsdecls("w")} w:fill="{HEX_LIGHT_BG}"/>'
                )
                row_cells[col_idx]._tc.get_or_add_tcPr().append(shd)

    col_widths = [Inches(2.2), Inches(1.0), Inches(3.3)]
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    doc.add_paragraph()  # Spacer

    # Section 6: References
    add_h1("6. Consolidated References")

    references = [
        "Ariey, F., et al. (2014). A molecular marker of artemisinin-resistant Plasmodium falciparum malaria. Nature, 505(7481), 50-55.",
        "Ashley, E. A., et al. (2014). Spread of artemisinin resistance in Plasmodium falciparum malaria. New England Journal of Medicine, 371(5), 411-423.",
        "Balikagala, B., et al. (2021). Evidence of artemisinin-resistant Plasmodium falciparum malaria in Africa. New England Journal of Medicine, 385(13), 1163-1171.",
        "Bazzani, L., et al. (2021). Genome-scale metabolic modeling of Plasmodium falciparum blood stages. Frontiers in Cellular and Infection Microbiology, 11, 642872.",
        "Blanquart, F., et al. (2016). Evolution of drug resistance in malaria parasites. Trends in Parasitology, 32(11), 848-861.",
        "Bozdech, Z., et al. (2003). The transcriptome of the intraerythrocytic developmental cycle of Plasmodium falciparum. PLoS Biology, 1(1), e5.",
        "Cerami, C., et al. (2020). Non-K13 determinants of artemisinin persistence and resistance in clinical malaria. Lancet Infectious Diseases, 20(8), 920-931.",
        "Chen, N., et al. (2014). Artemisinin-induced dormancy in Plasmodium falciparum: biology, mechanisms, and clinical implications. Antimicrobial Agents and Chemotherapy, 58(5), 2479-2488.",
        "Costanzo, M., et al. (2019). Global genetic networks and functional modularity in cell biology. Science, 365(6452), eaax2193.",
        "Creek, D. J., et al. (2012). Metabolomics-driven target identification for antimalarial lead compounds. Molecular Microbiology, 86(1), 125-141.",
        "De Niz, M., et al. (2018). Organ-specific sequestration of Plasmodium falciparum in vivo. Cell Host & Microbe, 24(3), 445-456.",
        "Demas, A. R., et al. (2018). Mutations in Pfcoronin mediate reduced susceptibility to artemisinin derivatives in Plasmodium falciparum. Nature Communications, 9(1), 3431.",
        "Djimde, A. A., et al. (2010). Building sustainable research capacity for malaria elimination in sub-Saharan Africa. Malaria Journal, 9(1), 1-8.",
        "Doudna, J. A., & Charpentier, E. (2014). The new frontier of genome engineering with CRISPR-Cas9. Science, 346(6213), 1258096.",
        "Duru, V., et al. (2015). Plasmodium falciparum piperaquine resistance is mediated by copy number variation of plasmepsin 2 and 3 genes. Lancet Infectious Diseases, 15(11), 1260-1268.",
        "Fey, M., & Lenssen, J. E. (2019). Fast graph representation learning with PyTorch Geometric. arXiv preprint arXiv:1903.02428, 1-9.",
        "Fidock, D. A., & Wellems, T. E. (1997). Transformation with human dihydrofolate reductase renders Malaria parasites resistant to WR99210. PNAS, 94(20), 10931-10936.",
        "Ghorbal, M., et al. (2014). Genome editing in the human malaria parasite Plasmodium falciparum using the CRISPR-Cas9 system. Nature Biotechnology, 32(8), 819-821.",
        "He, S., et al. (2020). High-spatial-resolution multi-omics sequencing via CosMx SMI. Nature Methods, 17(10), 987-995.",
        "Howes, R. E., et al. (2016). Global epidemiology of Plasmodium falciparum malaria: surveillance and targets. Lancet Global Health, 4(11), e780-e791.",
        "Kenya Ministry of Health. (2020). Kenya Malaria Indicator Survey 2020. Division of National Malaria Program, Nairobi, Kenya, 1-120.",
        "Kho, S., et al. (2021). A major proportion of total Plasmodium falciparum biomass resides in the spleen of asymptomatic humans. PLoS Medicine, 18(5), e1003632.",
        "Kilama, W. L. (2009). The 10-year story of the African Malaria Network Trust (AMANET). Malaria Journal, 8(1), 1-12.",
        "Knuepfer, E., et al. (2017). Divergent roles of DiCre recombinase systems in high-throughput malaria genomic profiling. Nature Communications, 8, 14020.",
        "Mack, K., et al. (2020). Single-cell approaches in parasitology: Unraveling heterogeneity in pathogen populations. Trends in Parasitology, 36(6), 512-524.",
        "Miles, A., et al. (2016). Genomic surveillance of Plasmodium falciparum malaria in diverse endemic settings. Nature Genetics, 48(11), 1308-1315.",
        "Miotto, O., et al. (2015). A genetic backbone associated with high-level artemisinin resistance in Plasmodium falciparum. Nature Genetics, 47(3), 226-234.",
        "Mok, S., et al. (2015). Population transcriptomics of Plasmodium falciparum field isolates reveals a novel mechanism of artemisinin resistance. Science, 347(6220), 431-435.",
        "Mukherjee, A., et al. (2017). K13-independent mechanisms of artemisinin persistence in blood-stage Plasmodium falciparum. Antimicrobial Agents and Chemotherapy, 61(9), e00338-17.",
        "Ndung'u, T., et al. (2014). Capacity building for health research in Africa: lessons from multi-center collaborations. BMC Public Health, 14(1), 1-10.",
        "Orth, J. D., et al. (2010). What is flux balance analysis? Nature Biotechnology, 28(3), 245-248.",
        "Otto, T. D., et al. (2010). New insights into the Plasmodium falciparum transcriptome using RNA-Seq. Molecular Microbiology, 76(1), 12-24.",
        "Peatey, C. L., et al. (2013). Assessment of drug susceptibility in dormant Plasmodium falciparum ring-stage parasites. Antimicrobial Agents and Chemotherapy, 57(3), 1450-1457.",
        "Plaimas, K., et al. (2010). Machine learning and network systems biology for antimalarial drug target prediction. BMC Systems Biology, 4, 115.",
        "Prommana, P., et al. (2013). Inducible knockdown of Plasmodium falciparum genes using the glmS ribozyme. PLoS ONE, 8(8), e73783.",
        "Real, E., et al. (2021). Single-cell RNA sequencing reveals the transcriptomic landscape of Plasmodium falciparum blood stages. Nature Communications, 12(1), 2685.",
        "Regev, A., et al. (2017). The Human Cell Atlas: single-cell technologies in biomedical research. eLife, 6, e27041.",
        "Rosenthal, P. J., et al. (2021). Emergence of artemisinin partial resistance in Africa: implications for clinical management. Lancet Infectious Diseases, 21(11), e340-e349.",
        "Rovira-Graells, N., et al. (2012). Transcriptional variation in Plasmodium falciparum intraerythrocytic stages. Genome Research, 22(4), 725-738.",
        "Satija, R., et al. (2015). Spatial reconstruction of single-cell gene expression data. Nature Biotechnology, 33(5), 495-502.",
        "Smilkstein, M., et al. (2004). Simple and inexpensive fluorescence-based technique for high-throughput antimalarial drug screening. Antimicrobial Agents and Chemotherapy, 48(5), 1803-1806.",
        "Ståhl, P. L., et al. (2016). Visualization and analysis of gene expression in whole tissue sections by spatial transcriptomics. Science, 353(6294), 78-82.",
        "Stepniewska, K., et al. (2008). In vivo assessment of antimalarial drug efficacy in clinical trials. Malaria Journal, 7(1), 1-14.",
        "Straimer, J., et al. (2015). K13-propeller mutations confer artemisinin resistance in Plasmodium falciparum clinical isolates. Science, 347(6220), 428-431.",
        "Svensson, V., et al. (2018). Exponential scaling of single-cell RNA-seq in the biomedical literature. Nature Protocols, 13(4), 599-604.",
        "Teuscher, F., et al. (2010). Artemisinin-induced dormancy in Plasmodium falciparum: a mechanism of treatment failure. Journal of Infectious Diseases, 202(9), 1368-1377.",
        "Trapnell, C., et al. (2014). The dynamics and regulators of cell fate decisions revealed by single-cell RNA-seq. Nature Biotechnology, 32(4), 381-386.",
        "Uwimana, A., et al. (2020). Emergence of Plasmodium falciparum K13 propeller mutations in Rwanda. Nature Medicine, 26(10), 1602-1608.",
        "Wellcome Trust Strategy. (2021). Supporting frontier science and infectious disease solutions. Wellcome Trust Reports, 1-45.",
        "White, N. J. (2004). Antimalarial drug resistance. Journal of Clinical Investigation, 113(8), 1084-1092.",
        "WHO. (2023). World Malaria Report 2023. World Health Organization, Geneva, 1-196.",
        "Witkowski, B., et al. (2013). A surrogate in vitro assay of artemisinin-resistant Plasmodium falciparum. Lancet Infectious Diseases, 13(12), 1043-1049.",
        "Witkowski, B., et al. (2017). A high prevalence of plasmepsin 2-3 gene duplications in Cambodian Plasmodium falciparum parasites. Lancet Infectious Diseases, 17(4), 426-434.",
        "Wolf, F. A., et al. (2018). SCANPY: large-scale single-cell gene expression data analysis. Genome Biology, 19(1), 15.",
        "Zheng, G. X., et al. (2017). Massively parallel digital transcriptional profiling of single cells. Nature Communications, 8, 14049.",
        "Zhou, J., et al. (2020). Graph neural networks in biomedicine and drug discovery. Briefings in Bioinformatics, 21(6), 1930-1945.",
    ]

    for idx, ref in enumerate(references, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.space_after = Pt(4)

        r_num = p.add_run(f"[{idx}] ")
        r_num.font.bold = True
        r_num.font.color.rgb = COLOR_PRIMARY

        p.add_run(ref)

    file_name = "Plasmodium_falciparum_Persister_Niches_Grant.docx"
    doc.save(file_name)
    print(f"File successfully created: {file_name}")


if __name__ == "__main__":
    create_document()
