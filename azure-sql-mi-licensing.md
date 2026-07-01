# Azure SQL Managed Instance Licensing

For Azure SQL Managed Instance (MI), licensing is much simpler than SQL Server on Azure VMs.

## Azure SQL MI Licensing Models

### 1. License Included

You pay for:

- vCores
- Storage
- Backup storage

SQL Server licensing is already included in the price. No existing SQL licenses required.

### 2. Azure Hybrid Benefit (AHB)

If the customer has:

- SQL Server Standard + Software Assurance (or subscription)
- SQL Server Enterprise + Software Assurance (or subscription)

they can apply those licenses to Azure SQL MI and pay the reduced base rate instead of the full license-included price.

## Conversion Rules

These are the important rules when you're looking at an Azure Migrate assessment.

### General Purpose Tier

| On-Prem License | Azure SQL MI GP |
| --- | ---: |
| 1 Standard Core | 1 vCore |
| 1 Enterprise Core | 4 vCores |

Enterprise customers receive the 4:1 virtualization benefit when targeting General Purpose.

### Example

Customer owns:

- SQL Enterprise 16 cores
- Active Software Assurance

They can cover:

- 64 vCores of General Purpose MI

16 x 4 = 64 vCores.

### Business Critical Tier

| On-Prem License | Azure SQL MI Business Critical |
| --- | ---: |
| 1 Standard Core | 0.25 vCore |
| 1 Enterprise Core | 1 vCore |

Business Critical does not receive the 4:1 multiplier.

### Example

Customer owns:

- SQL Enterprise 16 cores

They can cover:

- 16 Business Critical vCores

Not 64.
