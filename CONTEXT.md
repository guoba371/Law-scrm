# Law Customer Case Management

This context describes the customer and case-tracking language used by the local Excel-based law firm management tool.

## Language

**Customer**:
A person or organization appearing in the imported customer statistics workbook.
_Avoid_: Contact, lead account

**Customer Grade**:
The business status inferred from the imported row color: red means closed customer, yellow means valuable prospect, and light pink means prospect not closed.
_Avoid_: Color tag, customer type

**Closed Customer**:
A customer whose imported row is marked red and who should appear in the case tracking form.
_Avoid_: Deal, signed client

**Valuable Prospect**:
A customer whose imported row is marked yellow and may be worth future follow-up, but is not yet treated as a closed case customer.
_Avoid_: Hot lead

**Unclosed Prospect**:
A customer whose imported row is marked light pink and did not become a closed customer.
_Avoid_: Failed lead, invalid customer

**Case Tracking Index**:
The editable set of milestone fields attached to a closed customer, covering civil and criminal case progress.
_Avoid_: Case log, progress notes

**Open Case**:
A closed customer's case whose closure flag is not selected.
_Avoid_: Pending case, unfinished matter

**Closed Case**:
A closed customer's case whose closure flag is selected and is excluded from the open-case count.
_Avoid_: Completed matter
