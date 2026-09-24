# EPD Type III Submodel User Guide

## 1. Purpose of the Submodel
The EPD Type III Submodel represents Environmental Product Declarations (EPDs) within the Asset Administration Shell (AAS). Its purpose is to provide a programme-independent, machine-readable, semantic downstream exchange layer for LCA-based environmental data. 
It solves the problem of EPD fragmentation by harmonizing digital declarations into a standard AAS format. 
**Important**: This submodel is NOT an EPD publishing workflow (it does not replace program operators), nor is it a complete LCA project model (it does not contain the raw LCI background datasets). It is strictly for downstream consumption.

## 2. Intended users and use cases
This guide is intended for:
- **AAS platforms** hosting digital twins.
- **Digital Product Passport (DPP)** implementers tracking environmental impacts.
- **Customers** consuming automated B2B sustainability data.
- **Software developers** building EPD parsers.

## 3. Design principles
- **Programme Independence**: The submodel is designed to support multiple EPD programmes (e.g., PEP ecopassport, EPD International).
- **Semantics over Layout**: The structure represents semantic meaning, not the visual layout of a PDF.
- **Explicit Data**: Units and lifecycle stages are explicitly modeled rather than implied.
- **Optionality**: Optionality is driven by the fact that different Product Category Rules (PCRs) mandate different fields. 
- **Separation of Provenance**: Source-derived EPD data must be strictly segregated from illustrative or externally enriched example data.

## 4. High-level model overview
```text
EPD Type III [Submodel]
├── IdentificationPublication [SMC]
├── ProgramOperatorVerification [SMC]
├── ManufacturerProduct [SMC]
├── EPDScope [SMC]
├── FunctionalUnit [SMC]
├── DeclaredUnit [SMC]
├── ContentDeclaration [SMC]
├── LCAMethodology [SMC]
├── lca_StandardsPcr [SMC]
├── EnvironmentalResults [SML]
└── epdDocument [File]
```

## 5. Quick Start
### Minimum practical EPD instance
At a minimum, a useful EPD instance requires:
- `IdentificationPublication` (Registration numbers, Validity)
- `ManufacturerProduct` (Manufacturer name, Product name)
- `DeclaredUnit` or `FunctionalUnit`
- `EnvironmentalResults` (At least one indicator like GWP-total)

### Typical physical-product EPD
A physical product typically adds:
- `ManufacturingSites` (Where it was made)
- `ContentDeclaration` (Material weights and biogenic carbon)
- `EPDScope` (Whether it's a specific product, family, or average)

### Service / intangible EPD
For services (e.g., a software license EPD), `ContentDeclaration` and `ManufacturingSites` are generally omitted as they are physically inapplicable.

### Multi-product / family EPD
If the EPD covers a product family, the `EPDScope` describes the family, while the AAS Asset identifies the specific instantiated product variant covered by that family.

## 6. Section-by-section operating instructions

### IdentificationPublication
**Purpose**: Holds essential metadata regarding the EPD's publication and validity lifecycle.
**Element type**: SubmodelElementCollection.
**Use when**: Always (mandatory).
**Child elements**: Registration numbers, validity dates, document URLs.

### ProgramOperatorVerification
**Purpose**: Declares who published and verified the EPD.
**Element type**: SubmodelElementCollection.
**Use when**: Always (mandatory).
**Important**: Public verification statement URLs are not universally required and may be omitted if internal.

### ManufacturerProduct
**Purpose**: Describes the product and the organization declaring the EPD.
**Element type**: SubmodelElementCollection.
**Use when**: Always (mandatory).
**Important**: `manufacturerName` is the declaring entity, which may differ from `ManufacturingSite/siteOperator`.

### ManufacturingSites
**Purpose**: Represents physical manufacturing locations associated with the product system.
**Element type**: SubmodelElementList.
**Use when**: The EPD refers to identifiable manufacturing sites.
**May be omitted when**: No site info is available or appropriate (e.g. average/service declarations).
**List behaviour**: The list may contain multiple `ManufacturingSite` entries.
**Important distinctions**:
- `manufacturerName`: Declaring organization.
- `siteOperator`: Legal entity operating the specific physical site.
- `siteName`: Human-readable label.
- `SiteIdentifier`: Explicit, stable identity.
- Coordinates/LocationCode provide geographical matching evidence, not universal identity. Mobile sites may omit exact coordinates.

### EPDScope
**Purpose**: Defines the product boundary (e.g., Single Product, Average, Family).
**Element type**: SubmodelElementCollection.
**Use when**: Always.
**Important**: Differentiates between the generic EPD scope and the exact AAS Asset instance.

### FunctionalUnit / DeclaredUnit
**Purpose**: Defines the quantitative reference for the LCA results.
**Element type**: SubmodelElementCollection.
**Use when**: Depending on the PCR, use one or both. Functional Units represent performance over time, while Declared Units represent physical quantities (e.g., 1 kg). Both may coexist if explicitly required.

### ContentDeclaration
**Purpose**: Contains the material composition and biogenic carbon of the product and packaging.
**Element type**: SubmodelElementCollection.
**Use when**: Applicable for physical products where required by the PCR.
**Important**: This is NOT a complete chemical compliance BOM (e.g., REACH, RoHS), merely a summary for environmental analysis.

### LCAMethodology
**Purpose**: Describes the modeling assumptions (geography, time, energy models) behind the LCA.
**Element type**: SubmodelElementCollection.
**Use when**: Always (mandatory).

### lca_StandardsPcr
**Purpose**: Lists the PCRs, c-PCRs, and generic standards the EPD complies with.
**Element type**: SubmodelElementCollection.
**Use when**: Always (mandatory).

### EnvironmentalResults
**Purpose**: Contains the actual environmental impact and resource use indicators.
**Element type**: SubmodelElementList.
**Use when**: Always (mandatory).
**List behaviour**: Contains multiple `EnvironmentalResult` SMCs, each representing one indicator.

## 7. Complete field reference

| Model path | Element type | Cardinality / list behaviour | Datatype | Unit / coding scheme | Description | Example |
|---|---|---|---|---|---|---|
| `IdentificationPublication` | SMC | `1` | - | - | Metadata regarding the publication and validity. | - |
| `documentIds` | SML | `0..1` (multiple entries) | - | - | List of DocumentId SMCs representing registration numbers. | - |
| `programOperatorVersion` | Property | `0..1` | string | - | Internal version number assigned by the operator. | "1.0" |
| `dateOfIssue` | Property | `1` | date | - | Date the EPD was issued. | "2023-01-01" |
| `validUntil` | Property | `1` | date | - | Expiration date of the EPD. | "2028-01-01" |
| `languages` | SML | `0..1` (multiple entries) | - | - | ISO 639-1 language codes for the declaration. | "en" |
| `sourceDataFormat` | Property | `0..1` | string | ILCD/openEPD | Digital format of the original source data if available. | "ILCD" |
| `declarationUrl` | Property | `0..1` | string | URL | Public URL to retrieve the EPD PDF or dataset. | "https://..." |
| `ProgramOperatorVerification` | SMC | `1` | - | - | Details regarding the program operator and verifier. | - |
| `programOperator` | SMC | `1` | - | - | The organization running the EPD program. | - |
| `thirdPartyVerifier` | SMC | `0..1` | - | - | The independent verifier of the LCA. | - |
| `verifierAccreditationId` | Property | `0..1` | string | - | Accreditation number of the verifier. | "V-1234" |
| `epdDeveloper` | SMC | `1` | - | - | The person or organization that performed the LCA. | - |
| `verificationType` | Property | `1` | string | Internal/External | Whether verification was independent and external. | "External" |
| `verificationDate` | Property | `0..1` | date | - | Date the verification was completed. | "2023-01-01" |
| `ManufacturerProduct` | SMC | `1` | - | - | Product and manufacturer details. | - |
| `manufacturerName` | MultiLanguageProperty | `1` | langString | - | Name of the manufacturer associated with the product declared in the EPD. | "WAGO" |
| `ManufacturingSites` | SML | `0..1` (multiple entries) | - | - | List of physical manufacturing locations. | - |
| `productName` | MultiLanguageProperty | `1` | langString | - | Commercial name of the product. | "Terminal Block" |
| `productArticleNumberOfManufacturer` | Property | `0..1` | string | - | Primary SKU or article number. | "221-413" |
| `productDescription` | MultiLanguageProperty | `0..1` | langString | - | Brief text description of the product. | "3-conductor splicing connector" |
| `productImage` | File | `0..1` | file | - | Package-relative path to a product image. | "/aasx/image.png" |
| `massPerDeclaredUnit` | Property | `0..1` | decimal | kg | Mass of the product corresponding to the Declared Unit. | "0.05" |
| `referenceServiceLife` | Property | `0..1` | decimal | years | RSL of the product if a Functional Unit is used. | "20" |
| `applicableJurisdictions` | SML | `0..1` (multiple entries) | - | ISO 3166-1 | Countries where the product is sold/applicable. | "DE" |
| `manufacturingDescription` | MultiLanguageProperty | `0..1` | langString | - | Description of the manufacturing process. | "Injection molding..." |
| `EPDScope` | SMC | `1` | - | - | Defines the boundary of the represented product. | - |
| `scopeType` | Property | `1` | string | Single/Family... | The type of EPD representation. | "Family" |
| `coveredProductReferences` | SML | `0..1` (multiple entries) | - | - | List of specific SKUs or identifiers covered. | - |
| `scopeDescription` | MultiLanguageProperty | `0..1` | langString | - | Text describing the scope boundaries. | "Covers all colors." |
| `FunctionalUnit` | SMC | `0..1` | - | - | Performance-based reference unit. | - |
| `DeclaredUnit` | SMC | `0..1` | - | - | Physical quantity reference unit. | - |
| `quantity` | Property | `1` | decimal | - | Numerical amount. | "1.0" |
| `unit` | Property | `1` | string | - | The physical unit. | "kg" |
| `description` | MultiLanguageProperty | `0..1` | langString | - | Description of the unit. | "1 kg of steel" |
| `ContentDeclaration` | SMC | `0..1` | - | - | Material and biogenic carbon composition. | - |
| `productMass` | Property | `0..1` | decimal | kg | Total mass of the product. | "0.04" |
| `recycledContentPercentage` | Property | `0..1` | decimal | % | Share of recycled material. | "10.0" |
| `biogenicCarbonContent` | SMC | `0..1` | - | - | Container for biogenic carbon data. | - |
| `biogenicCarbonContentProduct` | Property | `0..1` | decimal | kg C | Biogenic carbon in the product itself. | "0.0" |
| `biogenicCarbonContentPackaging` | Property | `0..1` | decimal | kg C | Biogenic carbon in the packaging. | "0.01" |
| `LCAMethodology` | SMC | `1` | - | - | LCA assumptions and scope. | - |
| `lcaReferenceYear` | Property | `1` | integer | YYYY | The reference year for the LCA datasets. | "2023" |
| `timeRepresentativenessDescription` | Property | `0..1` | string | - | Text description of temporal validity. | "Data from 2022" |
| `geographicalScope` | Property | `1` | string | - | LCA representativeness (e.g., GLO, RER). | "GLO" |
| `geographicalDescription` | Property | `0..1` | string | - | Text description of geographical validity. | "Global supply chain" |
| `technologyDescription` | Property | `0..1` | string | - | Text description of technological validity. | "State of the art" |
| `lcaEnergyModel` | Property | `0..1` | string | - | Energy mix used in the LCA (e.g., DE, EU). | "DE" |
| `DeclaredStages` | SML | `0..1` (multiple entries) | - | - | List of stageCodes that are declared in this EPD. | "A1-A3" |
| `lca_StandardsPcr` | SMC | `1` | - | - | List of rules the EPD complies with. | - |
| `standardsCompliance` | SML | `0..1` (multiple entries) | - | - | Generic LCA standards (e.g. ISO 14025). | "ISO 14025" |
| `productCategoryRules` | SML | `0..1` (multiple entries) | - | - | Specific PCRs / PSRs applied. | "PEP-PCR-ed4" |
| `epdDocument` | File | `0..1` | file | PDF | Package-relative path to the EPD PDF document. | "/aasx/epd.pdf" |
| `EnvironmentalResults` | SML | `1` (multiple entries) | - | - | List of indicator results. | - |

*(Note: See Sections 8 and 13 for detailed nested structures of Environmental Results and Manufacturing Sites respectively).*

## 8. Environmental Results Guide
The `EnvironmentalResults` structure is an SML containing multiple `EnvironmentalResult` SMCs.
```text
EnvironmentalResults [SML]
└── EnvironmentalResult [SMC]
    ├── resultCategory [Property, 1] (e.g., "ImpactIndicator", "ResourceUse")
    ├── indicatorCode [Property, 1] (e.g., "GWP-total")
    ├── indicatorName [Property, 0..1]
    ├── unit [Property, 1] (Characterization unit, e.g., "kg CO2 eq")
    ├── unitId [Property, 0..1] (Semantic ID of the characterization unit)
    ├── baseUnit [SMC, 0..1] (Underlying physical/reference unit component)
    │   ├── unit [Property, 1]
    │   └── unitId [Property, 1]
    └── stageValues [SML, 1] (List of StageValue SMCs)
```
Each `EnvironmentalResult` specifies the `resultCategory` (e.g., `ImpactIndicator`, `ResourceUse`, `Waste`, `OutputFlow`) and an `indicatorCode` (e.g., `GWP-total`, `PENRT`, `HWD`). Note: illustrative indicators in examples do not necessarily originate from the actual published EPD.

### Environmental Indicator Registry
To ensure interoperability, `indicatorCode` and `unit` combinations **must** match the recognized combinations defined in the central `environmental-indicator-registry.yaml`. When a characterization unit represents a complex derived unit, its underlying physical base unit MUST be provided using the optional `baseUnit` SMC. Extensions beyond the registry must provide explicit semantic IDs.

## 9. Lifecycle / Stage Guide
Lifecycle stages are represented in the `stageValues` SML inside each indicator.
Each `StageValue` SMC contains a `stageCode` (e.g., `A1-A3`, `A4`, `B6`, `C4`, `D`, `Total`).
The model natively supports both distinct stages (A1, A2, A3) and aggregated stages (A1-A3, Upstream, Core). 
The `idShort` is normalized to remove hyphens (e.g., `stageCode = A1-A3` becomes `idShort = stage_A1_A3`).

## 10. Status semantics
Do NOT use numerical zero to represent "Not Declared".
`0.0 != NotDeclared != NotRelevant`
The `stageStatus` property correctly models this:
- `Declared`: The value is mathematically calculated (even if the value is 0.0).
- `NotDeclared`: The PCR does not require this stage, or it was omitted (value is absent).
- `NotRelevant`: The stage is physically impossible or irrelevant (value is absent).

## 11. Units and reference quantities
`FunctionalUnit` (performance-based, e.g., "lighting 10m2 for 1 year") and `DeclaredUnit` (physical quantity, e.g., "1 kg of steel") are distinct semantic concepts. While EPDs usually use one, they are not universally mutually exclusive and both may coexist depending on the PCR. Ensure units are explicit and machine-readable where possible.

## 12. EPD Scope and product families
The `EPDScope` describes the boundary. A Single Product EPD represents exactly one SKU. A Product Family EPD represents a range of SKUs. The AAS Asset represents the exact digital twin instance. Therefore, an AAS Asset representing "SKU-123" might link to an EPD whose scope is a "Family" covering "SKU-100 to SKU-199".

## 13. Manufacturing Sites
```text
ManufacturingSites [SML, 0..1]
└── ManufacturingSite [SMC, 0..*]
    ├── SiteIdentifiers [SML, 0..1]
    │   └── SiteIdentifier [SMC, 0..*]
    │       ├── identifierScheme [1]
    │       └── identifierValue [1]
    ├── siteName [Property, 0..1]
    ├── siteType [Property, 0..1] (Fixed / Mobile)
    ├── siteOperator [Property, 0..1]
    └── Location [SMC, 0..1]
        ├── countryCode [Property, 0..1]
        ├── region [Property, 0..1]
        ├── locality [Property, 0..1]
        ├── latitude [Property, 0..1] (Decimal)
        ├── longitude [Property, 0..1] (Decimal)
        └── LocationCodes [SML, 0..1]
            └── LocationCode [SMC, 0..*]
                ├── codingSystem [1]
                └── code [1]
```
- **manufacturerName vs siteOperator**: `manufacturerName` is the manufacturer associated with the product declared in the EPD. `siteOperator` is the legal entity physically operating the specific manufacturing site. When they differ, it explicitly models contract manufacturing.
- **SiteIdentifier**: The stable identity within a scheme (e.g., `identifierScheme = openEPDPlantID`, `identifierValue = 123`).
- **LocationCode**: Coded geographic matching evidence (e.g., `codingSystem = Open Location Code`).
- **Identity vs Matching**: Identity is established by `SiteIdentifier`. Coordinates and LocationCodes are geographical position/matching evidence, not universal identity. There is no normative "200m tolerance" rule.
- **Mobile Sites**: `siteType` = `Mobile`. Latitude/longitude are omitted; region-level location is used.
- **Provenance**: In examples, distinguishing SOURCE-DERIVED data from EXTERNALLY ENRICHED (e.g., highly precise coordinates) and ILLUSTRATIVE (synthetic examples) is strictly maintained.

## 14. Content Declaration
The `ContentDeclaration` summarizes material composition and biogenic carbon. It is applicable where PCRs demand it. It is NOT a complete chemical compliance BOM, REACH database, or general material master data.

## 15. LCA Methodology
This section documents LCA assumptions. The `geographicalScope` represents LCA validity (e.g., GLO), not the physical manufacturing site. The `lcaEnergyModel` documents the electricity grid mix assumed during manufacturing. If an existing property cannot be explained clearly, it is flagged as a `DOCUMENTATION-BLOCKING MODEL ISSUE`.

## 16. Verification
Details independent verification. Public verification-statement URLs are not universally required (many verifications are confidential/internal to the operator).

## 17. Programme / PCR / standards stack
- **Standard**: General LCA rules (e.g., ISO 14025, EN 15804).
- **PCR**: Product Category Rules governing the specific product category.
- **PSR**: Product Specific Rules (used in PEP ecopassport).

## 18. Documents and packaged files
Use package-relative paths (e.g., `/aasx/files/epd.pdf`) for embedded `epdDocument` or `productImage` files. **WARNING**: Never use absolute local file paths (e.g., `C:\Users\...`).

## 19. Semantic identifiers
- **ECLASS IRDIs**: Standardized semantic identifiers. Not generic URLs.
- **admin-shell.io**: Provisional working-group identifiers. These must not be presented as official IDTA allocations until explicitly authorized. 

## 20. Source-derived vs illustrative example data
- **SOURCE-DERIVED**: Value directly supported by the source EPD publication.
- **EXTERNALLY ENRICHED**: Value derived or added to demonstrate the model (e.g., highly precise geographic coordinates).
- **ILLUSTRATIVE**: Synthetic data created solely to demonstrate optional capabilities (e.g., a fictitious mobile plant).

## 21. Common modelling mistakes
- Using `0` instead of `NotDeclared`.
- Confusing FunctionalUnit and DeclaredUnit.
- Confusing `manufacturerName` (declaring organization) with `siteOperator` (physical operator).
- Using `siteName` or coordinates as stable canonical identity.
- Using `LocationCode` or `SiteIdentifier` without specifying the coding/identifier scheme.
- Hardcoding A1-D as fixed structural properties instead of using the flexible `stageValues` SML.
- Using absolute local file paths.
- Interpreting illustrative example values as actual claims about a manufacturer's product.

## 22. Programme interoperability
The model is structurally tested against supplied reference cases for:
- PEP ecopassport
- EPD International (non-construction)
- EPD International (construction / EN 15804)

## 23. Extension guidance
To extend the submodel:
- Prefer semantic extension over structural duplication.
- Avoid programme-specific top-level structures.
- Reuse existing list structures (SMLs).
- Use explicit identifier/code schemes.
- Retain backward compatibility where reasonable.
