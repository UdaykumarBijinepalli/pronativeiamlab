# Entra ID & M365 Identity Lifecycle Pipeline

Declarative Joiner-Mover-Leaver (JML) identity governance and RBAC automation pipeline built with Microsoft Graph REST API, PowerShell, Azure RBAC, and passwordless GitHub Actions OIDC.

## Architecture & Security Design

* **Git-Versioned Roster (oster/cohort01-trainees.csv):** Declarative HR/training source of truth defining EmployeeId, Department, JobTitle, Cohort, and Status (Active / Offboarded).
* **Zero-Secret Authentication (OIDC):** Authenticates GitHub Actions to Microsoft Entra ID via Workload Identity Federation—eliminating client secrets.
* **Least-Privilege Permissions:** Uses scoped Microsoft Graph permissions (Group.ReadWrite.All, User.ReadWrite.All, Organization.Read.All) and Resource-Group-scoped Azure RBAC—no tenant-wide directory admin roles.
* **Joiner / Mover (Active):** Synchronizes Entra user attributes (department, jobTitle), adds users to SG-Cohort01-Trainees, assigns available M365 license SKUs, and grants least-privilege Reader RBAC on g-cohort01-lab.
* **Leaver (Offboarded):** Updates directory status to Offboarded - Alumni, removes cohort Security Group membership, reclaims M365 licenses, and revokes Azure RBAC role assignments.
* **Compliance Audit Artifact:** Generates udit/identity-sync-report.json on every run and uploads it as a downloadable GitHub Actions artifact.
