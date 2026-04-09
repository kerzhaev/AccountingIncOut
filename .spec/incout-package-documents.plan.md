# IncOut Package Documents Plan

## Goal

Extend `AccountingIncOut` from a flat document register into a two-level model:

1. package/header record in `TableIncOut`
2. optional child documents linked to the package

The feature must preserve the current fast workflow while enabling:
- detailed package composition
- per-child-document 1C matching
- future outbound-letter integration with `CreateLetter`

## Current State

### AccountingIncOut

`TableIncOut` currently stores one logical record per row with these already relevant fields:
- service
- direction (`incoming` / `outgoing`)
- document type
- document number
- document amount
- FRP number/date
- counterparty
- executor
- service dispatch/return numbers and dates
- outgoing envelope/letter number and date (`columns 16/17`)
- execution mark
- confirmation status
- order info

The main UI is `UserFormVhIsh`.

### CreateLetter

`CreateLetter` already stores outgoing letter history in worksheet/table `Letters` with fields including:
- `Addressee`
- `OutgoingNumber`
- `OutgoingDate`
- `DocumentSum`
- `DocumentTypeKey`
- `Executor`
- row number DTO support through `clsLetterHistoryRecord`

This is enough to design future synchronization, but not enough for safe direct matching without a stable source identifier.

## Architectural Decision

### Package/Header Model

Keep `TableIncOut` as the parent package table.

One parent row represents a package or top-level document set such as:
- notice with attachments
- reconciliation act with supporting documents
- mixed incoming package

The parent row remains the main operational unit for:
- registration
- workflow state
- service dispatch / return tracking
- envelope / outgoing communication tracking

### Child Documents Model

Add a new child table, recommended name:
- `TableIncOutItems`

One child row represents one specific document inside a package, for example:
- invoice / waybill
- acceptance-transfer act
- notice
- reconciliation act
- other attachment document

Child rows are optional.

## Non-Goals For MVP

Do not implement in the first iteration:
- automatic migration of all historical rows into child documents
- automatic many-to-many synchronization with `CreateLetter`
- hard blocking on package/child amount mismatch
- complex tree nesting deeper than parent -> child
- rewriting current 1C matching in the same iteration

## Data Model

### Parent Table: `TableIncOut`

Keep current columns.

Add these new columns:

1. `PackageId`
- type: text
- required: yes for all new records
- purpose: stable unique identifier for the parent package

2. `PackageType`
- type: text/enum
- examples:
  - `notice_with_invoice`
  - `notice_with_act`
  - `act_only`
  - `invoice_only`
  - `mixed_package`

3. `AssetCategory`
- type: text/enum
- values:
  - `fixed_assets`
  - `inventory`
  - `mixed`

4. `DocumentStage`
- type: text/enum
- values:
  - `received`
  - `sent_to_service`
  - `returned_from_service`
  - `signed`
  - `posted_in_1c`
  - `confirmed`
  - `closed`

5. `CounterpartyNormalized`
- type: text
- purpose: normalized lookup field for future integrations and matching

6. `HasChildDocuments`
- type: boolean/text enum
- values: `True` / `False`

7. `ChildDocumentsCount`
- type: number

8. `ChildrenTotalAmount`
- type: number

9. `PackageAmountCheckStatus`
- type: text/enum
- values:
  - `not_checked`
  - `match`
  - `mismatch`

10. `Primary1COperationNumber`
- type: text

11. `Primary1COperationDate`
- type: date

12. `Primary1CMatchStatus`
- type: text/enum
- values:
  - `not_checked`
  - `exact`
  - `candidate`
  - `manual`
  - `not_found`

### Child Table: `TableIncOutItems`

Create a new worksheet-backed table.

Required columns:

1. `ItemId`
- stable unique child identifier

2. `PackageId`
- foreign key to parent package

3. `ItemOrder`
- order within package

4. `ItemDocumentType`
- internal type key

5. `ItemDocumentTypeDisplay`
- user-facing document type label

6. `ItemDocumentNumber`

7. `ItemDocumentDate`

8. `ItemAmount`

9. `CounterpartyName`
- snapshot inherited from parent

10. `CounterpartyNormalized`
- snapshot inherited from parent

11. `Direction`
- snapshot inherited from parent

12. `Service`
- snapshot inherited from parent

13. `Executor`
- snapshot inherited from parent

14. `OrderInfo`
- snapshot/inherited editable copy

15. `FRPNumber`
- snapshot inherited from parent

16. `FRPDate`
- snapshot inherited from parent

17. `ItemAssetCategory`
- `fixed_assets` / `inventory`

18. `ItemDescription`

19. `ItemQuantity`

20. `ItemUnit`

21. `BaseDocumentType`
- parent package base doc type snapshot

22. `BaseDocumentNumber`

23. `BaseDocumentDate`

24. `Matched1COperationNumber`

25. `Matched1COperationDate`

26. `Matched1CMatchStatus`
- values:
  - `not_checked`
  - `exact`
  - `candidate`
  - `manual`
  - `not_found`

27. `Matched1CConfidence`

28. `Matched1CComment`

29. `Matched1CMode`
- `auto`
- `manual`
- `grouped`

30. `IsPostedSeparately`
- boolean/text enum

31. `Notes`

32. `CreatedAt`

33. `UpdatedAt`

## UI Design

### Main Form: `UserFormVhIsh`

Keep it as the parent/package form.

Add the following UI elements:

1. package indicators
- `DocumentsInPackageCount`
- `ChildrenTotalAmount`
- `PackageAmountCheckStatus`

2. actions
- button `DocumentsInPackage...`
- optional button `AddChildDocument`

3. behavior
- existing fast save remains unchanged
- child documents are optional
- if child documents exist, package indicators update after save/close of the child form

### New Child Form: `UserFormPackageDocuments`

Purpose:
- manage all child documents linked to one parent package

Layout:

#### Header section (read-only package summary)
- `PackageId`
- base document type
- base document number
- base document date
- counterparty
- total package amount
- service
- executor

#### Child list section
Display columns:
- order
- document type
- number
- date
- amount
- asset category
- 1C status
- 1C operation number

Buttons:
- `Add`
- `Edit`
- `Delete`
- `Duplicate`
- `FillFromPackage`

#### Child edit section
Editable fields:
- document type
- document number
- document date
- amount
- asset category
- description
- quantity
- unit
- order info
- FRP number/date
- notes
- 1C operation number/date/status/comment

Buttons:
- `Save item`
- `Save and close`
- `Cancel`

## Data Entry Rules

### Fast mode
User saves only the parent package record.
No child rows are required.

### Detailed mode
User opens `DocumentsInPackage...` and adds one or more child rows.
Inherited values are prefilled from the parent package.

### Sum control
If child rows exist:
- `ChildrenTotalAmount = sum(ItemAmount)`
- compare with parent amount
- show warning on mismatch
- do not hard-block in MVP

## 1C Matching Strategy

### Parent package matching
Use only when:
- no child documents exist
- or user explicitly runs package-level matching

### Child document matching
Preferred path when child rows exist.

Use child fields for future exact matching:
- document type
- document number
- document date
- amount
- counterparty
- order info
- FRP

Rule:
- if children exist, child matching becomes primary
- parent status becomes aggregate/summary only

## Future CreateLetter Integration

The package model is the prerequisite for safe letter synchronization.

### Required future keys
In `AccountingIncOut`:
- `PackageId` on parent
- `ItemId` on child rows

In `CreateLetter` future integration layer:
- link letters to `PackageId`
- optionally link to concrete `ItemId` rows when one letter refers to specific documents inside the package

### Why this matters
A letter may cover:
- the whole package
- or only selected child documents

Direct matching by `counterparty + amount` is not reliable enough.
Stable IDs are required.

## Implementation Phases

### Phase 1 - Schema Foundation
1. add new parent columns to `TableIncOut`
2. add stable `PackageId` generation for new records
3. create worksheet/table `TableIncOutItems`
4. add repository/helpers for child CRUD

### Phase 2 - Parent Form Wiring
1. extend `UserFormVhIsh` with package indicators
2. add `DocumentsInPackage...` button
3. update parent save/load routines for new package columns
4. keep backward compatibility with records that have no children

### Phase 3 - Child Form MVP
1. create `UserFormPackageDocuments`
2. implement list/add/edit/delete for child rows
3. prefill inherited parent values
4. update package counters and child sum status

### Phase 4 - Validation and Usability
1. add parent/child amount comparison
2. add child ordering and duplication support
3. add clear status text for mismatch and missing child data
4. test mixed packages and single-child cases

### Phase 5 - Matching Readiness
1. add child-level 1C matching fields and display
2. define aggregate package status from child states
3. keep current package-level matching behavior untouched until dedicated refactor starts

## Verification Strategy

### Source validation
- exported modules remain the review surface
- `.frm`/`.frx` import/export cycle must stay stable

### Workbook validation through Excel COM
For each implementation phase, run:
1. import changed modules/forms into workbook
2. compile if available through VBE command surface
3. smoke open/close forms
4. create test parent package
5. create test child rows
6. verify package indicators and child count
7. verify saved values in worksheet tables
8. verify that legacy records without children still load normally

### Target smoke scenarios
1. fast package save without children
2. package save with two child documents
3. package amount mismatch warning
4. parent reopen after child edit
5. child delete updates parent counters
6. mixed assets package (`inventory` + `fixed_assets`)

## Risks

1. `TableIncOut` backward compatibility
- existing save/load code is tightly coupled to fixed columns

2. UserForm complexity growth
- `UserFormVhIsh` is already a hotspot
- new child form should absorb detail complexity instead of bloating the main form

3. Worksheet schema drift
- column creation and lookup must be deterministic

4. Matching ambiguity
- package-level and child-level matching must not conflict silently

## Recommended MVP Boundary

Implement first:
- `PackageId`
- parent metadata fields
- `TableIncOutItems`
- `UserFormPackageDocuments`
- parent-child save/load
- sum control warning

Do not implement yet:
- automatic historical migration
- full `CreateLetter` sync
- rewritten 1C matching engine

## Deliverables

1. parent schema extension in `TableIncOut`
2. new child table `TableIncOutItems`
3. package-document child form
4. updated parent form actions and indicators
5. COM smoke scripts / verification notes
6. follow-up plan for 1C child matching and `CreateLetter` synchronization
