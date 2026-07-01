# Windows Server Core License Conversion for Azure VMs

## Summary

If you're ignoring Azure Hybrid Benefit savings and just asking, **"How many Windows Server cores do I need to cover an Azure VM?"**, the conversion is essentially:

> **1 Windows Server core license per Azure vCPU, subject to Microsoft's minimums.**

For Azure Migrate business cases, I typically model it as:

```text
Required Windows Core Licenses = MAX(8, Azure VM vCPU Count)
```

per VM.

---

## Quick Rules

| Azure VM Size | Windows Server Core Licenses Required |
|---:|---:|
| 2 vCPU | 8 cores, minimum |
| 4 vCPU | 8 cores, minimum |
| 8 vCPU | 8 cores |
| 12 vCPU | 12 cores |
| 16 vCPU | 16 cores |
| 32 vCPU | 32 cores |
| 64 vCPU | 64 cores |

Microsoft requires a **minimum allocation of 8 Windows Server core licenses per Azure VM**, even if the VM is smaller than 8 vCPUs. Once you exceed 8 vCPUs, you license the actual VM core count.

---

## Examples

- **D2s_v5**, 2 vCPU → **8 Windows cores**
- **D4s_v5**, 4 vCPU → **8 Windows cores**
- **D8s_v5**, 8 vCPU → **8 Windows cores**
- **D16s_v5**, 16 vCPU → **16 Windows cores**
- **E32s_v5**, 32 vCPU → **32 Windows cores**

---

## Legacy Processor Licenses

If the customer still has old Windows Server processor licenses, Microsoft counts **one processor license as 16 core licenses** for Azure conversion purposes.

---

## Important Distinction: Windows Server vs. SQL Server

Do not confuse this with **SQL Server**, where there are 4:1 and other conversion ratios for some Azure services.

Windows Server does **not** have those multipliers.

For Windows Server VM licensing in Azure, think:

> **Azure VM vCPU count = Windows Server core count required, with a minimum of 8 cores per VM.**

---

## Standard vs. Datacenter: Same Azure Core Conversion

For Azure VM licensing, there is **no core conversion difference** between Windows Server Standard and Windows Server Datacenter.

The Azure licensing math is the same:

- 8-core minimum per VM
- Then 1 licensed core per Azure vCPU above 8
- A 16-vCPU VM requires 16 Windows Server cores whether you own Standard or Datacenter licenses

Where the editions differ is in the **usage rights**, not the conversion ratio.

| Area | Windows Server Standard | Windows Server Datacenter |
|---|---|---|
| Core conversion to Azure VM | Same | Same |
| Azure VM coverage | Same | Same |
| Concurrent on-prem use during migration | 180 days | Indefinite while licensed workloads remain on Azure |
| On-prem virtualization rights | Limited | Unlimited virtualization |

The practical implication for migration assessments is:

- If you're simply sizing Azure IaaS VMs, **Windows Server Standard and Datacenter both map cores the same way**.
- Datacenter becomes valuable when the customer has heavily virtualized on-prem hosts because a fully licensed Datacenter host allows unlimited Windows Server VMs on that host.
- Once those workloads are moved into Azure IaaS VMs, the Azure VM core requirement itself is not discounted because the customer owns Datacenter instead of Standard.

For Azure Migrate/TCO conversations, I usually tell customers:

> For Azure VM licensing, ignore Standard vs. Datacenter when calculating required cores. Count the Azure VM vCPUs, subject to the 8-core minimum. The edition matters primarily because of the rights and economics on the source environment, not because of a different Azure core conversion ratio.

---

## The Nuance: Existing License Pools

One nuance: if you're trying to calculate **how many Azure VMs can be covered by a pool of existing Windows licenses**, then Standard and Datacenter can differ because of their licensing rights and reassignment rules.

But the **per-VM Azure core calculation remains the same**.

There are really two different questions:

1. **How many licenses do I need for a specific Azure VM?**
2. **How many Azure VMs can I cover with my existing license pool?**

---

## Question 1: How Many Licenses Do I Need for a Specific Azure VM?

For this question, **Standard and Datacenter are identical**.

| Azure VM | Required Windows Core Licenses |
|---:|---:|
| 4 vCPU VM | 8 cores, minimum |
| 8 vCPU VM | 8 cores |
| 16 vCPU VM | 16 cores |
| 32 vCPU VM | 32 cores |

The math is the same regardless of whether those licenses are Windows Server Standard or Datacenter.

---

## Question 2: How Many Azure VMs Can I Cover with My Existing License Pool?

This is where Standard and Datacenter diverge.

Think of a customer with a 16-core physical server on-prem.

### Windows Server Standard

A fully licensed 16-core Standard server only grants rights for **up to two OSEs/VMs** on that host.

Example:

- 16 physical cores licensed with Standard
- Running 10 VMs on-prem

To legally run those 10 VMs, the customer must stack Standard licenses multiple times.

A rough illustration:

- First 16-core license set = 2 VMs
- Second 16-core license set = 2 more VMs
- Third 16-core license set = 2 more VMs
- Etc.

So many customers with dense VMware environments end up licensing the same 16-core host several times.

### Windows Server Datacenter

A fully licensed 16-core Datacenter host gets:

- Unlimited Windows Server VMs on that host

Whether they run:

- 10 VMs
- 50 VMs
- 100 VMs

They still only license the host once.

That is why highly virtualized VMware customers almost always buy Datacenter.

---

## Why This Matters for Azure Migrations

Suppose a customer has:

- 10 VMware VMs
- Each VM will become an 8-vCPU Azure VM

Each Azure VM requires 8 cores to cover it.

So Azure requires:

```text
10 VMs × 8 cores = 80 Windows Server cores
```

It does **not matter** that those VMs originated from a Datacenter-licensed host.

Azure looks at the cores being assigned to the Azure VMs, not how efficiently the customer licensed the source VMware cluster.

---

## Practical Customer Conversation

When you're doing Azure Migrate/TCO discussions, I usually think about it this way:

### On-prem

- **Standard** = pay for VM rights
- **Datacenter** = pay for host rights

### Azure

- Both editions become a pool of licensable cores
- Coverage is assigned VM-by-VM based on Azure vCPU count
- The "unlimited virtualization" benefit of Datacenter largely disappears because Azure VMs are licensed individually

---

## Customer Example

Customer has:

- VMware cluster
- 4 hosts
- 32 cores per host
- 100 Windows VMs

### On-prem

- Standard would be extremely expensive because the customer would repeatedly stack licenses.
- Datacenter is usually the obvious choice due to unlimited virtualization.

### After migration to Azure

- What matters is the size of the Azure VMs.
- A 16-vCPU Azure VM consumes 16 Windows cores.
- Ten 16-vCPU Azure VMs consume 160 Windows cores.
- Whether those licenses originated as Standard or Datacenter does not change the Azure VM coverage math.

---

## Key Takeaway

For Azure Native migrations, the key takeaway is:

> **Datacenter's value comes from unlimited virtualization on-prem. Once you're calculating Azure VM coverage, the metric becomes total licensable cores, and the per-VM conversion is the same for Standard and Datacenter.**

This is why VMware-heavy customers often look "overlicensed" with Datacenter on-prem, yet can still come up short on Windows core coverage once hundreds of Azure VMs are sized and counted individually.
