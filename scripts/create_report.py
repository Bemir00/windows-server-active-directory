"""Generate ActiveDirectory_Project_Report.docx from scratch using zipfile."""
import zipfile, os
from xml.sax.saxutils import escape

OUT = 'ActiveDirectory_Project_Report.docx'

# ── XML helpers ────────────────────────────────────────────────────────────────

def rpr(bold=False, size=24, color=None, font=None, italic=False, mono=False):
    parts = []
    if font or mono:
        fn = "Courier New" if mono else (font or "Calibri")
        parts.append(f'<w:rFonts w:ascii="{fn}" w:hAnsi="{fn}" w:cs="{fn}"/>')
    if bold:   parts.append('<w:b/><w:bCs/>')
    if italic: parts.append('<w:i/><w:iCs/>')
    if color:  parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    return '<w:rPr>' + ''.join(parts) + '</w:rPr>' if parts else ''

def run(text, bold=False, size=24, color=None, italic=False, mono=False):
    rp = rpr(bold=bold, size=size, color=color, italic=italic, mono=mono)
    return f'<w:r>{rp}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def para(content_xml, style='Normal', align=None, before=0, after=160, line=None, indent_left=0):
    ppr_parts = [f'<w:pStyle w:val="{style}"/>']
    sp = f'<w:before w:val="{before}"/><w:after w:val="{after}"/>'
    if line: sp += f'<w:line w:val="{line}" w:lineRule="auto"/>'
    ppr_parts.append(f'<w:spacing {sp.replace("<w:spacing ",""[:-1])}')  # rebuild
    if align: ppr_parts.append(f'<w:jc w:val="{align}"/>')
    if indent_left: ppr_parts.append(f'<w:ind w:left="{indent_left}"/>')
    spacing_xml = f'<w:spacing w:before="{before}" w:after="{after}"'
    if line: spacing_xml += f' w:line="{line}" w:lineRule="auto"'
    spacing_xml += '/>'
    align_xml = f'<w:jc w:val="{align}"/>' if align else ''
    ind_xml = f'<w:ind w:left="{indent_left}"/>' if indent_left else ''
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>{spacing_xml}{align_xml}{ind_xml}</w:pPr>'
    return f'<w:p>{ppr}{content_xml}</w:p>'

def heading1(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            f'<w:spacing w:before="360" w:after="120"/></w:pPr>'
            f'{run(text, bold=True, size=32, color="1F3864")}</w:p>')

def heading2(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="Heading2"/>'
            f'<w:spacing w:before="240" w:after="80"/></w:pPr>'
            f'{run(text, bold=True, size=26, color="2E5090")}</w:p>')

def p(text, bold=False, size=24, align=None, before=60, after=120):
    return para(run(text, bold=bold, size=size), align=align, before=before, after=after)

def p2(parts_list, before=60, after=120):
    """parts_list = list of (text, bold, italic, mono)"""
    runs = ''
    for item in parts_list:
        if isinstance(item, str):
            runs += run(item)
        else:
            text, bold, italic, mono = (item + (False, False, False))[:4] if not isinstance(item, tuple) else (item + (False,) * (4 - len(item)))
            runs += run(text, bold=bold, italic=italic, mono=mono)
    return para(runs, before=before, after=after)

def code_block(text):
    """Monospace code block with grey background simulation (indent)."""
    lines = text.split('\n')
    result = ''
    for line in lines:
        r = f'<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">{escape(line) if line else " "}</w:t></w:r>'
        result += f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="0" w:after="0"/><w:ind w:left="720"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>{r}</w:p>'
    return result

def page_break():
    return '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'

def empty_para():
    return '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="0" w:after="80"/></w:pPr></w:p>'

def bullet(text, bold=False):
    rp = rpr(bold=bold, size=24)
    r = f'<w:r>{rp}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            f'<w:spacing w:before="40" w:after="40"/></w:pPr>{r}</w:p>')

def toc_entry(num, title, page_hint):
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="60" w:after="60"/>'
            f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs></w:pPr>'
            f'{run(f"{num}.  {title}", bold=False, size=22)}'
            f'<w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
            f'<w:tab/><w:t>{page_hint}</w:t></w:r></w:p>')

# ── Document body ──────────────────────────────────────────────────────────────

def build_body():
    body = []

    # ── TITLE PAGE ────────────────────────────────────────────────────────────
    body.append(empty_para())
    body.append(empty_para())
    body.append(empty_para())
    body.append(para(run('Engineering of Windows Server OS 2026', bold=True, size=52, color='1F3864'), align='center', before=0, after=200))
    body.append(para(run('Active Directory Project Report', bold=False, size=36, color='2E5090'), align='center', before=0, after=400))
    body.append(para(run('Course Project', bold=False, size=28, color='444444'), align='center', before=0, after=80))
    body.append(para(run('Domain: mydomain.com  |  Platform: VMware Workstation  |  OS: Windows Server 2019', size=22, color='666666'), align='center', before=0, after=80))
    body.append(para(run('2026', size=24, color='444444'), align='center', before=0, after=80))
    body.append(page_break())

    # ── TABLE OF CONTENTS ────────────────────────────────────────────────────
    body.append(para(run('Table of Contents', bold=True, size=36, color='1F3864'), align='center', before=0, after=240))
    toc_items = [
        (1, 'Introduction to Windows Server', 3),
        (2, 'Introduction to Active Directory', 5),
        (3, 'User Account Configuration (GUI and PowerShell)', 8),
        (4, 'Groups Configuration', 11),
        (5, 'Configuring File and Folder Permissions', 14),
        (6, 'Using Group Policies to Remotely Configure Computers', 17),
        (7, 'Using Group Policies to Remotely Install Software', 20),
        (8, 'Using and Configuring DNS Service', 23),
        ('', 'Conclusion', 26),
    ]
    for num, title, pg in toc_items:
        body.append(toc_entry(num, title, pg))
    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 1 — Introduction to Windows Server
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('1. Introduction to Windows Server'))

    body.append(p('Windows Server is a group of enterprise-class server operating systems developed by Microsoft, designed to share services with multiple users and provide extensive administrative control of data storage, applications, and corporate networks. Unlike consumer editions of Windows, Windows Server is optimized to function as the backbone of an organization\'s IT infrastructure, providing services such as file and print sharing, remote desktop hosting, web services, identity management, and much more. Windows Server runs in data centers, cloud environments, and small business offices around the world, making it one of the most widely deployed server platforms in existence today.'))

    body.append(heading2('1.1 Versions and Installation Options'))
    body.append(p('Microsoft has released several versions of Windows Server over the years, with Windows Server 2019 and Windows Server 2022 being among the most current at the time of writing. For this lab project, we used Windows Server 2019 Standard Edition, which is a full-featured version that includes a graphical desktop environment (also called Desktop Experience). Microsoft also offers a Server Core installation option — a minimal interface version without a traditional graphical interface, designed to reduce the attack surface, lower resource usage, and simplify maintenance. Windows Server 2022 introduced enhanced security features, including secured-core server capabilities, improved HTTPS and TLS 1.3 support, and SMB over QUIC, making it an excellent choice for modern enterprise deployments. Regardless of the version, all Windows Server editions from 2019 onward are fully compatible with the Active Directory concepts covered in this report.'))

    body.append(heading2('1.2 Key Server Roles'))
    body.append(p('Windows Server offers a wide array of server roles that can be installed and configured through the Server Manager. Key roles include:'))
    body.append(bullet('Active Directory Domain Services (AD DS) — manages centralized user authentication, authorization, and directory services for the entire domain.'))
    body.append(bullet('DNS Server — resolves hostnames to IP addresses and is tightly integrated with AD DS; Active Directory cannot function without a properly configured DNS server.'))
    body.append(bullet('DHCP Server — automatically assigns IP addresses, subnet masks, gateways, and DNS server information to network clients, eliminating the need for manual IP configuration.'))
    body.append(bullet('File and Storage Services — provides shared storage, file server capabilities, and tools for managing disk volumes and storage spaces.'))
    body.append(bullet('Remote Access Services (RAS) / Routing and Remote Access (RRAS) — allows the server to act as a router, VPN server, or NAT gateway, enabling clients on an internal network to access the internet through the server.'))
    body.append(bullet('Web Server (IIS) — hosts web applications, intranet portals, and websites within the organization.'))

    body.append(heading2('1.3 Why Organizations Use Windows Server'))
    body.append(p('Organizations choose Windows Server for many compelling reasons. It integrates seamlessly with other Microsoft products such as Microsoft 365, Azure Active Directory, Exchange Server, and SQL Server. It provides centralized management of thousands of user accounts, computers, and devices through Active Directory, all from a single administrative console. Its Group Policy feature allows administrators to enforce security settings, software deployments, and configurations across an entire organization from one location. Furthermore, Windows Server includes robust security features — including Windows Defender Antivirus, Windows Firewall with Advanced Security, BitLocker Drive Encryption, and Role-Based Access Control (RBAC) — making it suitable for highly regulated industries such as healthcare, finance, and government. The familiarity of the Windows interface also reduces training overhead for administrators already familiar with desktop Windows.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 2 — Introduction to Active Directory
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('2. Introduction to Active Directory'))

    body.append(p('Active Directory (AD) is a directory service developed by Microsoft for Windows domain networks. It was first introduced with Windows Server 2000 and has since become one of the most critical components of enterprise IT infrastructure worldwide. Active Directory stores information about network objects — such as users, computers, groups, printers, and other resources — and makes this information available to administrators and users in a structured and searchable manner. It acts as the central authority for authentication and authorization in a Windows domain environment: when a user logs into a domain-joined computer, Active Directory verifies their identity and grants them access to the resources they are authorized to use. This centralized identity management is what makes Active Directory indispensable in organizations of any size.'))

    body.append(heading2('2.1 Key Concepts and Terminology'))
    body.append(p('To understand Active Directory, it is essential to understand its fundamental building blocks:'))
    body.append(bullet('Domain — the fundamental unit of organization in Active Directory. A domain is a collection of objects (users, computers, groups) that share a common database, security policies, and a namespace. In our lab, the domain is mydomain.com.'))
    body.append(bullet('Forest — the highest-level container in Active Directory, representing the complete AD instance. A forest can contain one or more domains that share a common schema and global catalog.'))
    body.append(bullet('Tree — a group of domains sharing a contiguous namespace (e.g., mydomain.com and hr.mydomain.com form a tree within the same forest).'))
    body.append(bullet('Domain Controller (DC) — a server running Active Directory Domain Services that is responsible for authenticating users, storing the AD database, and enforcing security policies. In our lab, our Windows Server 2019 VM serves as the DC.'))
    body.append(bullet('Organizational Unit (OU) — a container within a domain used to organize objects logically (e.g., by department or location) and to apply Group Policy settings to specific subsets of users or computers.'))
    body.append(bullet('LDAP (Lightweight Directory Access Protocol) — the protocol used to query and modify the Active Directory database. Tools like Active Directory Users and Computers communicate with AD via LDAP.'))
    body.append(bullet('Kerberos — the primary authentication protocol used by Active Directory. It uses tickets issued by a Key Distribution Center (KDC, which runs on the DC) to prove identity without transmitting passwords over the network.'))

    body.append(heading2('2.2 Workgroup vs. Domain'))
    body.append(p('Before Active Directory, most small networks used Workgroup configurations. In a Workgroup, each computer manages its own local accounts and security independently — there is no centralized authentication. Every user needs a separate account created on each individual computer they want to access, and each machine has its own separate password database. As networks grow beyond a handful of computers, this becomes completely unmanageable. A Domain solves this fundamental problem by centralizing user account management on the Domain Controller. A single domain account allows a user to authenticate once and access all resources they are authorized for across the entire network — a concept known as Single Sign-On (SSO). One account creation, one password policy, one place to disable an account when someone leaves the organization.'))

    body.append(heading2('2.3 AD DS vs. Azure Active Directory'))
    body.append(p('Microsoft also offers Azure Active Directory (Azure AD, now rebranded as Microsoft Entra ID), a cloud-based identity service that extends traditional Active Directory capabilities to cloud and hybrid environments. While on-premises AD DS focuses on managing resources within a local network and uses Kerberos and LDAP protocols, Azure AD is designed for managing identities in cloud applications such as Microsoft 365, Azure services, and third-party SaaS applications, using modern protocols like OAuth 2.0 and OpenID Connect. Organizations with both on-premises and cloud resources can synchronize their on-premises AD with Azure AD using Azure AD Connect, creating a hybrid identity solution where users have one identity across both environments. For this course project, we focused exclusively on the traditional on-premises Active Directory Domain Services.'))

    body.append(heading2('2.4 Lab Environment — Creating the mydomain.com Domain'))
    body.append(p('In our lab environment, we built a fully functional Active Directory domain using VMware Workstation with two virtual machines. The Domain Controller (DC) runs Windows Server 2019 and was configured with two network adapters: a NAT adapter (named INTERNET) that receives a DHCP-assigned IP address from the host router for internet connectivity, and an internal adapter (named INTERNAL_X) configured with the static IP address 172.16.0.1, a subnet mask of 255.255.255.0, no default gateway (since the DC itself is the gateway), and the DNS server set to 127.0.0.1 (loopback — so the server queries its own DNS service). The client machine (CLIENT1) runs Windows 10 and is connected exclusively to the VMnet0 internal network, obtaining all network configuration from the DC via DHCP.'))
    body.append(p('The domain was created by first installing the Active Directory Domain Services role through Server Manager (Add Roles and Features wizard). After installation, the Server Manager dashboard displayed a notification prompting promotion of the server to a domain controller. Running the Active Directory Domain Services Configuration Wizard, we selected "Add a new forest" and entered the Root domain name as mydomain.com. After configuring the Directory Services Restore Mode (DSRM) password and completing the wizard, the server automatically restarted. Upon reboot, the login screen displayed MYDOMAIN\\Administrator, confirming the domain had been successfully created. Simultaneously, the DNS Server role was automatically installed and configured with the appropriate zones for mydomain.com.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 3 — User Account Configuration
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('3. User Account Configuration (GUI and PowerShell)'))

    body.append(p('User accounts in Active Directory represent individuals — or services — that need to access network resources. Every person who logs onto a domain-joined computer or accesses a domain resource requires a domain user account. These accounts are stored in the Active Directory database on the Domain Controller and contain a rich set of attributes: the user\'s full name, logon name, password (stored as a hash), email address, phone number, department, manager, group memberships, logon hours, and more. Centralized management of user accounts is one of the primary administrative responsibilities in an Active Directory environment and one of the most significant advantages of the domain model over Workgroups.'))

    body.append(heading2('3.1 Creating Users Through the GUI (ADUC)'))
    body.append(p('The primary graphical tool for managing user accounts in Active Directory is Active Directory Users and Computers (ADUC), accessible from Server Manager > Tools > Active Directory Users and Computers. To create a new user, an administrator navigates to the appropriate Organizational Unit (OU) in the left-hand tree, right-clicks and selects New > User. The New Object – User wizard opens and prompts for the user\'s first name, last name, initials, and User Logon Name (the account name used to authenticate against the domain, e.g., jsmith@mydomain.com). After clicking Next, the administrator sets an initial password and chooses from options such as "User must change password at next logon" (recommended for security), "Password never expires" (use cautiously), "Account is disabled," and "User cannot change password." The account is created immediately and is available for authentication.'))

    body.append(heading2('3.2 Dedicated Domain Administrator Account'))
    body.append(p('In our lab, one of the first tasks after creating the domain was to create a dedicated domain admin account instead of continuing to use the built-in Administrator account for daily administrative work. This is a fundamental security best practice. The built-in Administrator account is a well-known, high-value target for attackers — its name is predictable, and if compromised, it provides complete control over the domain. By creating a named admin account (following a convention such as a-firstname.lastname to distinguish admin accounts from regular accounts), we maintain accountability through audit logs: every action taken is attributed to a specific named administrator. The dedicated account was created in ADUC and then added to the Domain Admins group, granting it full administrative control over the entire domain. Subsequently, all administrative tasks were performed using this account.'))

    body.append(heading2('3.3 Batch User Creation with PowerShell'))
    body.append(p('For environments with tens, hundreds, or thousands of users, manually creating accounts through the GUI is impractical and error-prone. PowerShell provides a powerful and efficient solution through the ActiveDirectory module\'s New-ADUser cmdlet. In our lab, we wrote a PowerShell script that reads a text file containing a list of first and last names (one name per line), constructs usernames following the pattern [first initial][last name] (e.g., John Smith becomes jsmith), and creates all accounts automatically with a default password. The script also creates a dedicated Organizational Unit (_USERS) to house all the generated accounts.'))

    body.append(p('The following PowerShell script demonstrates the batch user creation process used in our lab:'))
    body.append(code_block(
r"""# -----------------------------------------------------------------------
# Batch User Creation Script for Active Directory
# Creates domain users from a list of names in names.txt
# -----------------------------------------------------------------------

$PASSWORD_FOR_USERS = $env:AD_DEFAULT_PASSWORD
$USER_FIRST_LAST_LIST = Get-Content .\names.txt

# Convert plain-text password to SecureString
$password = ConvertTo-SecureString $PASSWORD_FOR_USERS -AsPlainText -Force

# Create Organizational Unit to hold new users
New-ADOrganizationalUnit -Name _USERS -ProtectedFromAccidentalDeletion $False

foreach ($n in $USER_FIRST_LAST_LIST) {
    $first    = $n.Split(" ")[0].ToLower()
    $last     = $n.Split(" ")[1].ToLower()
    $username = "$($first.Substring(0,1))$($last)".ToLower()

    Write-Host "Creating user: $($username)" -ForegroundColor Cyan

    New-ADUser `
        -AccountPassword      $password `
        -GivenName            $first `
        -Surname              $last `
        -DisplayName          $username `
        -Name                 $username `
        -EmployeeID           $username `
        -PasswordNeverExpires $true `
        -Path "ou=_USERS,$(([ADSI]'').distinguishedName)" `
        -Enabled $true
}

Write-Host "User creation complete!" -ForegroundColor Green"""))

    body.append(p('The script uses the Get-Content cmdlet to read the names file, splits each line into first and last name components, builds the username by concatenating the first initial with the last name, and calls New-ADUser with all required parameters. The -Path parameter places each user in the _USERS OU using the ADSI provider to dynamically retrieve the domain\'s distinguished name. This approach scales to thousands of users with no additional effort beyond maintaining the names list.'))

    body.append(heading2('3.4 Benefits of PowerShell Automation'))
    body.append(p('The advantages of PowerShell automation for user management extend far beyond simple time savings. Scripts are repeatable and consistent — every account is created with the exact same settings, eliminating human error from manual entry. They are auditable — the script itself serves as documentation of what was done and how. They are extensible — additional parameters such as department, manager, office location, and phone number can be added to the New-ADUser call with minimal changes. PowerShell can also be used for bulk modifications (resetting passwords, disabling accounts of departed employees, moving users between OUs), compliance reporting (generating lists of inactive or unlicensed accounts), and integration with HR systems through CSV imports, enabling automated provisioning and de-provisioning workflows.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 4 — Groups Configuration
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('4. Groups Configuration'))

    body.append(p('Groups in Active Directory are container objects that hold user accounts, computer accounts, and other groups as members. The fundamental purpose of groups is to simplify permission and policy management: instead of assigning permissions to individual user accounts — which becomes unmanageable as organizations grow — administrators assign permissions to groups and then control access by managing group membership. When a new employee joins and needs access to certain resources, the administrator adds them to the appropriate groups, and they immediately inherit all associated permissions. When an employee leaves, removing them from all groups (or simply disabling their account) instantly revokes all their access.'))

    body.append(heading2('4.1 Group Types'))
    body.append(p('Active Directory has two fundamental types of groups, which serve different purposes:'))
    body.append(bullet('Security Groups — used to assign permissions to resources such as files, folders, printers, and applications. Security groups can also be mail-enabled for email distribution. These are the groups used for access control in an Active Directory environment.'))
    body.append(bullet('Distribution Groups — used exclusively for email distribution lists (e.g., a group email address for the entire Finance department) and cannot be used to assign permissions to resources. Distribution groups require an email system such as Microsoft Exchange to be useful.'))

    body.append(heading2('4.2 Group Scopes'))
    body.append(p('Groups also have different scopes that determine who can be a member and where the group can be used to assign permissions. Understanding scopes is important for designing a scalable and manageable permission structure:'))
    body.append(bullet('Domain Local — can contain members from any domain in the forest (users, computers, and groups from any domain), but can only be used to assign permissions within its own domain. Typically used for granting access to specific local resources (e.g., a shared folder on a specific server).'))
    body.append(bullet('Global — can only contain members from the same domain, but can be assigned permissions in any domain within the forest. Typically used to organize users with similar roles or functions (e.g., all users in the Sales department).'))
    body.append(bullet('Universal — can contain members from any domain in the forest and can be assigned permissions in any domain. Ideal for multi-domain forests, but membership changes are replicated to all Global Catalog servers, so they should be used judiciously.'))
    body.append(p('A recommended best practice for permission management is the AGDLP strategy: place user Accounts into Global groups based on their role, add Global groups into Domain Local groups based on the resource being accessed, and assign Permissions to the Domain Local groups. This provides flexibility and clarity in permission management.'))

    body.append(heading2('4.3 Creating and Managing Groups via GUI'))
    body.append(p('Creating a group through Active Directory Users and Computers is straightforward. An administrator right-clicks on the target OU (e.g., a "Groups" OU), selects New > Group, enters a descriptive group name (e.g., "IT_Department" or "HR_ReadOnly"), selects the appropriate Group Scope and Group Type, and clicks OK. The group is created and visible in ADUC. To add members, the administrator double-clicks the group, navigates to the Members tab, and clicks Add to search for and add user accounts or other groups. The Member Of tab shows which other groups this group belongs to, enabling nested group structures.'))

    body.append(heading2('4.4 Managing Groups with PowerShell'))
    body.append(p('PowerShell provides the New-ADGroup and Add-ADGroupMember cmdlets for efficient group management. To create a new security group and populate it with members from our lab environment, the following commands would be used:'))
    body.append(code_block(
r"""# Create a new Global Security Group for the IT Department
New-ADGroup -Name "IT_Department" `
            -GroupCategory Security `
            -GroupScope Global `
            -Path "OU=Groups,DC=mydomain,DC=com" `
            -Description "IT Department staff with elevated access"

# Create a Domain Local group for file share permissions
New-ADGroup -Name "FileShare_HR_ReadOnly" `
            -GroupCategory Security `
            -GroupScope DomainLocal `
            -Path "OU=Groups,DC=mydomain,DC=com"

# Add users to the IT_Department group
Add-ADGroupMember -Identity "IT_Department" -Members jsmith, bjones, alee

# Add IT_Department global group to the file share domain local group
Add-ADGroupMember -Identity "FileShare_HR_ReadOnly" -Members "IT_Department"

# List all members of a group
Get-ADGroupMember -Identity "IT_Department" | Select-Object Name, SamAccountName"""))

    body.append(heading2('4.5 Practical Use Cases'))
    body.append(p('In our lab domain (mydomain.com), groups provide a logical and manageable structure for access control. Practical examples of groups include: IT_Admins (Global Security group containing all IT administrators, added to the local Administrators group on workstations via Group Policy), HR_Department (Global Security group for HR staff, granted Read/Write to the HR shared folder), Students_All (Global Security group containing all student accounts, used to apply specific logon restrictions via Group Policy), Managers (Global Security group for management-level users who need access to financial reports), and Software_Deploy_Targets (a group of computer accounts used to target software deployment GPOs). This structure ensures that as people join, leave, or change roles within the organization, access can be adjusted simply by modifying group membership rather than touching individual permissions across dozens of resources.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 5 — File and Folder Permissions
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('5. Configuring File and Folder Permissions'))

    body.append(p('One of the most important functions of a Windows Server domain is providing centralized, controlled access to files and folders stored on the server. Windows uses two complementary layers of permissions for network file sharing: NTFS permissions, which are enforced at the file system level and apply to both local and network access, and Share permissions, which apply only when a resource is accessed over the network. Understanding both layers and how they interact is essential for correctly securing shared resources in an Active Directory environment.'))

    body.append(heading2('5.1 NTFS Permissions vs. Share Permissions'))
    body.append(p('NTFS (New Technology File System) permissions are set on files and folders directly on the disk and are enforced by the operating system regardless of how the resource is accessed — locally or over the network. They provide granular control and support a full permission hierarchy. Share permissions, on the other hand, are applied when a folder is shared over the network and form the outer boundary of access control. When both Share and NTFS permissions are in place simultaneously, the effective permission a user has when accessing a folder over the network is the more restrictive of the two. For this reason, a common administrative best practice is to set Share permissions to Full Control for the "Everyone" group (or the specific security group), and use NTFS permissions exclusively for fine-grained access control, as NTFS permissions offer far more flexibility and precision.'))

    body.append(heading2('5.2 Sharing Folders Over the Network'))
    body.append(p('To share a folder on the network, an administrator right-clicks the folder in Windows Explorer, selects Properties, and navigates to the Sharing tab. Clicking Advanced Sharing reveals options to enable sharing, set the Share Name (the name visible to network users when they browse the server), limit the number of simultaneous users, and configure Share permissions. Alternatively, the File and Storage Services role in Server Manager provides a centralized management interface for all shared folders on the server, supporting both SMB (Windows file sharing) and NFS (for Linux/Unix clients) shares. Once shared, the folder becomes accessible via the UNC (Universal Naming Convention) path \\\\ServerName\\ShareName or \\\\IPAddress\\ShareName.'))

    body.append(heading2('5.3 NTFS Permission Levels'))
    body.append(p('NTFS permissions are organized into five standard permission levels for folders (listed from most to least permissive):'))
    body.append(bullet('Full Control — the user can read, write, modify, execute, delete, change attributes, change permissions, and take ownership of the folder and its contents.'))
    body.append(bullet('Modify — the user can read, write, execute, and delete files and subfolders, but cannot change permissions or take ownership.'))
    body.append(bullet('Read & Execute — the user can list folder contents, view file attributes, read file contents, and run executable files.'))
    body.append(bullet('Read — the user can list folder contents and view file and folder attributes, but cannot run executables or write any changes.'))
    body.append(bullet('Write — the user can create new files and subfolders and write data to existing files, but cannot read existing files without Read permission.'))
    body.append(p('Permissions can be explicitly assigned (directly applied) or inherited from a parent folder. By default, permissions flow down from parent to child folders through inheritance. An administrator can break inheritance on a specific subfolder to apply custom permissions that differ from the parent.'))

    body.append(heading2('5.4 Using Groups for Permission Management'))
    body.append(p('Using Active Directory security groups to manage folder permissions is a fundamental best practice. Instead of assigning permissions directly to individual user accounts — which requires updating folder ACLs every time a user joins, leaves, or changes roles — an administrator creates appropriately named security groups (e.g., HR_ReadOnly, Finance_FullAccess, IT_Admins), grants those groups the appropriate NTFS permissions on the relevant folders, and then simply manages group membership to control access. In our lab, this approach means that when a new member of the HR team is hired, the administrator adds their account to the HR_Department group, and they automatically receive the correct access to all HR resources without touching any folder permissions.'))

    body.append(heading2('5.5 UNC Paths and Network Drive Mapping'))
    body.append(p('Network shared folders are accessed using Universal Naming Convention (UNC) paths in the format \\\\ServerName\\ShareName. For example, a shared folder named "HR_Files" on our Domain Controller would be accessed as \\\\DC\\HR_Files or \\\\172.16.0.1\\HR_Files. Users can access shared folders directly by typing the UNC path in Windows Explorer\'s address bar or in the Run dialog (Windows + R). For convenience, network drives can be mapped to UNC paths, making a shared folder appear as a local drive letter (e.g., H: for the HR share). Drive mapping can be done manually through File Explorer > Map Network Drive, or automatically for all users in a department through Group Policy (discussed in Section 6), ensuring users always have the appropriate drives available when they log in.'))

    body.append(heading2('5.6 Best Practices — Principle of Least Privilege'))
    body.append(p('The principle of least privilege is the cornerstone of secure permission management: users and groups should be granted only the minimum permissions necessary to perform their legitimate job functions, and no more. For example, most employees only need Read access to shared company policy documents — granting Full Control or Modify to all users is an unnecessary security risk. Practical best practices include: never granting the Everyone group broad access, using dedicated security groups rather than individual user assignments, regularly auditing folder permissions to remove stale access rights, documenting the purpose of each shared resource and its permission assignments, and using NTFS permissions as the primary access control mechanism rather than relying on Share permissions alone. These practices significantly reduce the risk of accidental data exposure or malicious insider threats.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 6 — Group Policies: Remote Configuration
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('6. Using Group Policies to Remotely Configure Computers and Users'))

    body.append(p('Group Policy is one of the most powerful and widely used features of Active Directory, enabling administrators to centrally manage and configure operating system settings, security policies, software, and user environments across all domain-joined computers and user accounts — without physically touching each machine. A Group Policy Object (GPO) is a collection of settings that defines what a system will look like and how it will behave for a defined group of users or computers. GPOs are created and managed using the Group Policy Management Console (GPMC), which provides a unified interface for creating, editing, linking, and reporting on Group Policy across the domain.'))

    body.append(heading2('6.1 How Group Policy Works — LSDOU Processing Order'))
    body.append(p('Understanding the order in which Group Policy is applied is critical for predicting the effective result when multiple GPOs with potentially conflicting settings exist. Policies are applied in the following sequence, known as LSDOU: Local policy (settings configured directly on the local computer), then Site policy (GPOs linked to the Active Directory site the computer belongs to), then Domain policy (GPOs linked to the domain), and finally Organizational Unit policy (GPOs linked to the OU containing the user or computer, with parent OUs processed before child OUs). When multiple GPOs specify conflicting values for the same setting, the last one applied wins — meaning OU-level policies override domain-level, which override site-level, which override local. Administrators can modify this behavior using enforcement (forcing a GPO to win regardless of processing order) or block inheritance (preventing child OUs from inheriting GPOs from parent OUs or the domain).'))

    body.append(heading2('6.2 Computer Configuration Policies'))
    body.append(p('Computer Configuration settings within a GPO apply to the computer itself, regardless of which user is logged on, and take effect when the computer starts up. Common computer configuration policies in enterprise environments include:'))
    body.append(bullet('Password Policy — enforcing minimum password length (e.g., 8 characters), complexity requirements (upper/lowercase, numbers, symbols), and maximum password age through Computer Configuration > Windows Settings > Security Settings > Account Policies > Password Policy.'))
    body.append(bullet('Account Lockout Policy — automatically locking an account after a defined number of failed login attempts (e.g., lock after 5 failures for 30 minutes) to protect against brute-force attacks.'))
    body.append(bullet('Screen Lock / Interactive Logon — requiring users to lock their screen after a period of inactivity and requiring Ctrl+Alt+Del for logon to prevent spoofing attacks.'))
    body.append(bullet('Removable Storage Restrictions — denying read or write access to USB storage devices to prevent data theft or malware introduction.'))
    body.append(bullet('Windows Update Configuration — controlling when and how Windows updates are downloaded and installed, directing computers to a WSUS server for centrally managed patching.'))

    body.append(heading2('6.3 User Configuration Policies'))
    body.append(p('User Configuration settings apply to the user, regardless of which computer they log on to, and take effect at logon. Common user configuration policies include:'))
    body.append(bullet('Desktop Wallpaper — enforcing a standard corporate wallpaper across all workstations (User Configuration > Administrative Templates > Desktop > Desktop Wallpaper), preventing users from changing it to maintain a professional appearance.'))
    body.append(bullet('Mapped Network Drives — automatically connecting users to their department\'s network share (User Configuration > Windows Settings > Drive Maps), eliminating the need for users to configure drive mappings manually.'))
    body.append(bullet('Folder Redirection — redirecting special folders like Documents, Desktop, and AppData to a network location on the server, ensuring data is automatically backed up and accessible from any computer.'))
    body.append(bullet('Internet Browser Settings — configuring the default browser homepage, proxy settings, and restricting access to certain browser configuration options.'))
    body.append(bullet('Software Restrictions / AppLocker — defining which applications users are permitted to run, preventing the execution of unauthorized or potentially malicious software.'))

    body.append(heading2('6.4 Creating and Linking a GPO'))
    body.append(p('To create a new GPO, an administrator opens Group Policy Management (GPMC) from Server Manager > Tools, right-clicks on the target domain or OU in the left-hand tree, and selects "Create a GPO in this domain, and Link it here." After providing a descriptive name for the GPO, the administrator right-clicks the new GPO and selects Edit to open the Group Policy Management Editor. Settings are organized in a tree structure under Computer Configuration and User Configuration, each containing Policies (administrative templates and security settings enforced by the system) and Preferences (settings that act as defaults but can be modified by users). After configuring the desired settings, the editor is closed and the GPO is immediately active for the linked scope.'))

    body.append(heading2('6.5 Forcing Policy Refresh — gpupdate /force'))
    body.append(p('Group Policy is refreshed automatically on a schedule: every 90 minutes for regular computers and users (with a randomized offset to avoid network storms), every 5 minutes for domain controllers, and always at computer startup and user logon. However, during administration and testing, waiting 90 minutes is impractical. The command gpupdate /force, run in an elevated Command Prompt or PowerShell on the target machine, forces an immediate refresh of all Group Policy settings — both Computer and User configuration — re-applying every policy regardless of whether it has changed. This is invaluable during troubleshooting and initial GPO deployment. The /force flag ensures all policies are re-applied even if they haven\'t changed since the last refresh.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 7 — Group Policies: Software Installation
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('7. Using Group Policies to Remotely Install and Configure Software'))

    body.append(p('Beyond configuring operating system settings, Group Policy can be used to automatically install, update, or remove software on domain-joined computers — all without requiring manual intervention or physical access to each machine. Software deployment through Group Policy is managed under Software Settings within either Computer Configuration or User Configuration in a GPO. When a computer starts up or a user logs on, Windows checks the applicable GPOs and automatically installs any software packages specified in those GPOs. This capability is particularly valuable for ensuring that all domain computers have required software installed — security agents, productivity tools, or department-specific applications — consistently and without administrative overhead.'))

    body.append(heading2('7.1 Technical Requirements for Software Deployment'))
    body.append(p('Software deployment via Group Policy has specific technical requirements that must be met for the process to work correctly:'))
    body.append(bullet('MSI Format — the software package must be in MSI (Microsoft Installer) format. MSI is the standardized Windows package format that supports automated, silent installations. EXE-based installers cannot be deployed directly through Group Policy software installation; they must first be repackaged into MSI format using tools like the Microsoft ORCA editor or third-party repackaging tools.'))
    body.append(bullet('Network Share — the MSI file must be stored on a network share accessible via UNC path (e.g., \\\\DC\\Software\\app.msi). The share must be readable by the computer accounts (not user accounts) during startup, before any user is logged on, since computer-targeted installations happen in the computer context at boot time.'))
    body.append(bullet('Permissions — the network share must grant Read access to the Domain Computers group (or the specific computer accounts in scope) to allow machines to download the package during startup. The SYSTEM account on each target computer must be able to reach the share.'))

    body.append(heading2('7.2 Assigned vs. Published Deployment'))
    body.append(p('Group Policy software deployment supports two deployment methods:'))
    body.append(bullet('Assigned (mandatory) — the software is automatically installed without user interaction. When assigned to computers, installation happens at startup; when assigned to users, the software is advertised (appears in the Start menu) and installs automatically when the user first opens it or logs on. This method is used for required software that all targeted users or computers must have (e.g., antivirus, VPN client, corporate productivity suite).'))
    body.append(bullet('Published (optional) — only available for user-targeted deployment. The software appears in Control Panel > Programs > Get Programs from the network, allowing users to install it on demand. The installation is not automatic — users choose when (or whether) to install it. This is suitable for optional tools that only some users need, such as specialized applications for specific job functions.'))

    body.append(heading2('7.3 Step-by-Step Software Deployment via GPO'))
    body.append(p('The following process describes how to deploy an MSI application to all workstations in the domain using a computer-targeted GPO:'))
    body.append(bullet('Step 1 — Create the software share: Create a folder on the Domain Controller (e.g., C:\\Software), share it with the name "Software," and verify that the Domain Computers group has Read access to the share.'))
    body.append(bullet('Step 2 — Copy the MSI: Place the MSI file (e.g., application.msi) in the shared folder, accessible via \\\\DC\\Software\\application.msi.'))
    body.append(bullet('Step 3 — Create the GPO: In Group Policy Management, create a new GPO linked to the OU containing the target computers (e.g., the Workstations OU or the domain root for all computers).'))
    body.append(bullet('Step 4 — Configure software installation: Edit the GPO, navigate to Computer Configuration > Policies > Software Settings > Software Installation, right-click, select New > Package, browse to the UNC path of the MSI file (\\\\DC\\Software\\application.msi), and choose the deployment method (Assigned).'))
    body.append(bullet('Step 5 — Apply and test: Close the editor. On the target client (CLIENT1), run gpupdate /force or restart the computer. The software will be installed automatically during the next startup.'))

    body.append(heading2('7.4 Limitations and Enterprise Alternatives'))
    body.append(p('While Group Policy software deployment is a practical solution for small to medium environments, it has notable limitations. It only supports MSI-based packages without wrapping. It provides no progress reporting or installation success/failure tracking beyond the Windows Event Log. It does not support patch management or software updates well. It has limited targeting capabilities (essentially OU-based). For enterprise-scale software management with comprehensive reporting, inventory, patch management, and compliance tracking, Microsoft System Center Configuration Manager (SCCM, now called Microsoft Endpoint Configuration Manager) and Microsoft Intune (for cloud-managed and modern devices) provide far more sophisticated capabilities. However, for the scope of this lab and many real-world small organization scenarios, GPO-based software deployment remains a free, built-in, and entirely adequate solution.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 8 — DNS Service
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('8. Using and Configuring DNS Service'))

    body.append(p('The Domain Name System (DNS) is a hierarchical, distributed naming system that translates human-readable domain names (such as mydomain.com or www.example.com) into numerical IP addresses that computers use to communicate across networks. DNS is often described as the "phone book of the internet." In an Active Directory environment, DNS is not merely important — it is absolutely fundamental. Active Directory depends entirely on DNS for its operation: clients use DNS to locate Domain Controllers for authentication, Domain Controllers use DNS to find each other for replication, and domain-joined computers use DNS to resolve the names of servers, shared folders, and applications. Without a correctly configured DNS server, an Active Directory domain cannot function.'))

    body.append(heading2('8.1 Automatic DNS Installation with AD DS'))
    body.append(p('When we installed Active Directory Domain Services on our Windows Server 2019 virtual machine and ran the AD DS Configuration Wizard to promote it to a Domain Controller, the DNS Server role was automatically installed and configured alongside it. This is by design: the AD DS promotion wizard detects that DNS is required, verifies whether a DNS server is already present and properly configured, and if not, offers (and recommends) installing and configuring DNS automatically. Once installed, the DNS server immediately created the mydomain.com forward lookup zone and began registering all the resource records that Active Directory requires — most importantly, SRV (Service Location) records that advertise the location of the Domain Controller\'s LDAP, Kerberos, and other AD services to all domain clients.'))

    body.append(heading2('8.2 DNS Zones — Forward and Reverse Lookup'))
    body.append(p('DNS manages two primary types of zones that serve different resolution purposes:'))
    body.append(bullet('Forward Lookup Zones — the most commonly used zone type. It maps hostnames and domain names to IP addresses. For example, the mydomain.com forward lookup zone resolves DC.mydomain.com to 172.16.0.1. When CLIENT1 wants to find the Domain Controller or any other resource by name, it queries the forward lookup zone. In the DNS Manager on our DC, the mydomain.com zone contains A records for the DC itself and SRV records for all AD services.'))
    body.append(bullet('Reverse Lookup Zones — maps IP addresses back to hostnames (also called PTR or pointer records). For example, querying for 172.16.0.1 returns DC.mydomain.com. Reverse lookup is used by diagnostics tools (like nslookup), certain authentication systems, and email servers that verify sender identity. While not strictly required for basic AD functionality, reverse lookup zones are recommended for complete DNS operation and troubleshooting capability.'))

    body.append(heading2('8.3 DNS Record Types'))
    body.append(p('DNS uses several types of resource records to answer different kinds of queries:'))
    body.append(bullet('A (Address) Record — maps a hostname to an IPv4 address. The most fundamental record type in DNS, used for resolving server names. Example: DC.mydomain.com → 172.16.0.1.'))
    body.append(bullet('AAAA Record — maps a hostname to an IPv6 address, the IPv6 equivalent of an A record.'))
    body.append(bullet('CNAME (Canonical Name) Record — creates an alias from one hostname to another. For example, "fileserver.mydomain.com" could be an alias for "DC.mydomain.com," allowing the fileserver name to resolve to the DC\'s IP without a separate A record.'))
    body.append(bullet('MX (Mail Exchanger) Record — specifies the mail server(s) responsible for accepting email for a domain. Critical for email delivery in environments with Exchange Server.'))
    body.append(bullet('SRV (Service Location) Record — perhaps the most important record type for Active Directory. SRV records advertise the location of specific network services. Active Directory registers SRV records for _ldap._tcp.mydomain.com (pointing clients to the LDAP port on the DC) and _kerberos._tcp.mydomain.com (pointing clients to the Kerberos authentication service), among many others. Domain-joined computers use these SRV records to automatically locate Domain Controllers and AD services without any manual configuration.'))
    body.append(bullet('PTR (Pointer) Record — used in reverse lookup zones to map an IP address back to a hostname.'))

    body.append(heading2('8.4 Why the DC Uses 127.0.0.1 as Its DNS Server'))
    body.append(p('A critical and often misunderstood configuration detail in our lab is that the Domain Controller\'s internal network interface (INTERNAL_X, IP 172.16.0.1) was configured to use 127.0.0.1 as its Preferred DNS Server — the loopback address, which always refers to "this computer itself." This configuration is correct and intentional. Since the Domain Controller is its own DNS server (it hosts the DNS role and the mydomain.com zone), it should resolve DNS queries by querying itself. If we had pointed the DC to an external DNS server such as 8.8.8.8 (Google\'s public DNS), the DC would be unable to resolve names within the mydomain.com zone (which only exists on our local DNS server), and Active Directory would fail to function correctly. The loopback address ensures the DC always uses its own locally hosted DNS service, where all AD-critical records are stored.'))

    body.append(heading2('8.5 CLIENT1 DNS Configuration and Verification'))
    body.append(p('CLIENT1 receives its DNS server configuration automatically from the DHCP service on the Domain Controller. When CLIENT1 requests an IP lease from DHCP, the DC\'s DHCP server provides: an IP address (172.16.0.100, the first address in the configured scope of 172.16.0.100–200), a subnet mask of 255.255.255.0, a default gateway of 172.16.0.1 (through which internet traffic is routed via RAS/NAT), and the DNS server address of 172.16.0.1 (the DC itself). This means CLIENT1 sends all DNS queries to the Domain Controller, which can resolve both mydomain.com internal names and external internet names (by forwarding queries to upstream DNS servers such as 8.8.8.8).'))
    body.append(p('After joining the domain and obtaining the IP address via DHCP, we verified DNS operation on CLIENT1 using two commands. Running ipconfig confirmed that the DNS Suffix Connection-Specific was mydomain.com, the IPv4 Address was 172.16.0.100, the Subnet Mask was 255.255.255.0, and the Default Gateway was 172.16.0.1 — all exactly as configured. Running ping mydomain.com confirmed that the name mydomain.com successfully resolved to 172.16.0.1 with replies received and 0% packet loss, proving that DNS name resolution was fully operational and that CLIENT1 could communicate with the Domain Controller by name.'))

    body.append(heading2('8.6 DNS Scavenging and Dynamic Updates'))
    body.append(p('In a dynamic environment where computers join and leave the network, or where DHCP assigns different IP addresses to computers over time, DNS records can become stale — outdated records pointing to IP addresses that no longer belong to a particular computer. Dynamic DNS (DDNS) is the mechanism that allows clients and DHCP servers to automatically register and update their DNS records as IP addresses change. In our lab, when CLIENT1 receives an IP address from DHCP, the DHCP server (or CLIENT1 itself, depending on configuration) automatically creates or updates the A record for CLIENT1.mydomain.com in the DNS zone.'))
    body.append(p('DNS scavenging is a complementary maintenance feature that automatically removes stale resource records that have not been refreshed within a configurable interval. In the DNS Manager, scavenging is configured at both the server level and the zone level by setting the No-refresh interval (how long after registration a record cannot be refreshed) and the Refresh interval (how long after the no-refresh interval before the record is considered stale and eligible for deletion). Enabling dynamic updates and scavenging keeps the DNS zone accurate and prevents the accumulation of phantom records that can cause name resolution failures and confuse administrators during troubleshooting.'))

    body.append(page_break())

    # ═══════════════════════════════════════════════════════════════════════════
    # CONCLUSION
    # ═══════════════════════════════════════════════════════════════════════════
    body.append(heading1('Conclusion'))

    body.append(p('This project provided comprehensive hands-on experience with Windows Server and Active Directory — the technologies that form the backbone of enterprise Windows networking in organizations worldwide. Beginning with the physical and virtual infrastructure, we built a complete Active Directory environment from scratch: configuring a Domain Controller with dual network interfaces (NAT for internet connectivity and an internal VMnet0 network for domain communication), installing and promoting Active Directory Domain Services to create the mydomain.com domain, and configuring supporting services including DNS, DHCP, and RAS/NAT to provide a fully functional network environment for domain-joined clients.'))
    body.append(p('The user and group management sections demonstrated the power of centralized identity management — both through the graphical Active Directory Users and Computers interface and through PowerShell automation. The ability to create hundreds of user accounts through a brief script, and to control access to all resources simply by managing group membership, represents one of the most significant productivity advantages that Active Directory offers administrators. Our CLIENT1 Windows 10 virtual machine successfully joined the mydomain.com domain and was assigned network configuration (IP 172.16.0.100, gateway 172.16.0.1, DNS 172.16.0.1) through DHCP, with successful ping resolution of the domain confirming full functionality.'))
    body.append(p('Group Policy demonstrated how an Active Directory administrator can enforce configurations, security settings, and software deployments across an entire organization from a single management console. From mandatory password complexity policies to automatic software installation and drive mapping, Group Policy eliminates the need to configure each machine individually and ensures consistency and compliance across the entire domain. The importance of the LSDOU processing order and the gpupdate /force command for immediate policy application were key practical takeaways from this section.'))
    body.append(p('The DNS section revealed the foundational role that DNS plays in making Active Directory function — a role that is often invisible but always critical. The configuration choices made in our lab (loopback DNS on the DC\'s internal NIC, DC as the DNS server for all DHCP clients, automatic DNS installation during AD promotion) are not arbitrary but reflect carefully designed dependencies between these services. Without correct DNS configuration, every aspect of Active Directory — authentication, replication, group policy application, and resource access — would fail.'))
    body.append(p('The skills and concepts covered in this project — Windows Server administration, Active Directory domain management, user and group configuration, file permission management, Group Policy, and DNS — constitute the core competency set for roles such as Systems Administrator, Network Administrator, Active Directory Engineer, and IT Support Engineer. These topics also form the foundation for more advanced areas of study, including Active Directory Federation Services (ADFS) for federated identity, Azure Active Directory for hybrid cloud identity, Microsoft Endpoint Manager (Intune + SCCM) for modern device management, and Active Directory Certificate Services (AD CS) for public key infrastructure. This lab has provided a solid and practical foundation upon which to build expertise in enterprise Windows infrastructure.'))

    return '\n'.join(body)


# ── Styles XML ─────────────────────────────────────────────────────────────────

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="w14">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="160"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="360" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/><w:bCs/>
      <w:color w:val="1F3864"/>
      <w:sz w:val="32"/><w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="80"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/><w:bCs/>
      <w:color w:val="2E5090"/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
</w:styles>'''

NUMBERING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>'''

SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml" Id="rId4"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

FOOTER_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Normal"/>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="0"/>
    </w:pPr>
    <w:r>
      <w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="888888"/></w:rPr>
      <w:t xml:space="preserve">Active Directory Project Report  |  mydomain.com  |  Page </w:t>
    </w:r>
    <w:fldSimple w:instr=" PAGE ">
      <w:r>
        <w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="888888"/></w:rPr>
        <w:t>1</w:t>
      </w:r>
    </w:fldSimple>
  </w:p>
</w:ftr>'''

CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Active Directory Project Report</dc:title>
  <dc:subject>Engineering of Windows Server OS 2026</dc:subject>
  <dc:description>Course project report covering Active Directory setup and administration</dc:description>
  <cp:lastModifiedBy>Claude</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-05-26T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-26T00:00:00Z</dcterms:modified>
</cp:coreProperties>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>'''

# ── Assemble and write the DOCX ────────────────────────────────────────────────

body_content = build_body()

# Section properties with footer reference
sectPr = '''<w:sectPr>
  <w:footerReference w:type="default" r:id="rId4"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>'''

document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
<w:body>
{body_content}
{sectPr}
</w:body>
</w:document>'''

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml', CONTENT_TYPES_XML)
    zf.writestr('_rels/.rels', ROOT_RELS)
    zf.writestr('word/document.xml', document_xml)
    zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
    zf.writestr('word/styles.xml', STYLES_XML)
    zf.writestr('word/numbering.xml', NUMBERING_XML)
    zf.writestr('word/settings.xml', SETTINGS_XML)
    zf.writestr('word/footer1.xml', FOOTER_XML)
    zf.writestr('docProps/core.xml', CORE_XML)
    zf.writestr('docProps/app.xml', APP_XML)

print(f'Created: {OUT}')
print(f'Size: {os.path.getsize(OUT):,} bytes')
