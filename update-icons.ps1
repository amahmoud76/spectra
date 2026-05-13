# Replace Font Awesome icons with SVGs

# EditPanel.tsx - Replace replace file icon  
$content = Get-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Raw
$content = $content -replace '<i className="fa-regular fa-arrow-left-arrow-right" aria-hidden="true" />', '<img src={require(''../../assets/icons/replace.svg'')} alt="" style={{ width: ''1em'', height: ''1em'', display: ''inline-block'' }} aria-hidden="true" />'
Set-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Value $content
Write-Host 'Updated replace file icon'

# EditPanel.tsx - Replace re-activate icon
$content = Get-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Raw
$content = $content -replace '<i className="fa-regular fa-arrow-rotate-left" aria-hidden="true" />', '<img src={require(''../../assets/icons/re-activate.svg'')} alt="" style={{ width: ''1em'', height: ''1em'', display: ''inline-block'' }} aria-hidden="true" />'
Set-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Value $content
Write-Host 'Updated re-activate icon'

# EditPanel.tsx - Replace delete icon in quickActions
$content = Get-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Raw
$content = $content -replace '<i className="fa-regular fa-trash" aria-hidden="true" />\s*<span>Delete</span>', '<img src={require(''../../assets/icons/delete.svg'')} alt="" style={{ width: ''1em'', height: ''1em'', display: ''inline-block'' }} aria-hidden="true" />`n                <span>Delete</span>'
Set-Content 'src\webparts\spectra\components\EditPanel\EditPanel.tsx' -Value $content
Write-Host 'Updated delete icon'

# SPECTRA.tsx - Add circle-exclamation icon beside "No document found"
$content = Get-Content 'src\webparts\spectra\components\SPECTRA.tsx' -Raw
$content = $content -replace '<div className={styles.noResultsHint}>No document found</div>', '<div className={styles.noResultsHint}><img src={require(''../assets/icons/circle-exclamation.svg'')} alt="" style={{ width: ''1em'', height: ''1em'', display: ''inline-block'', marginRight: ''6px'' }} aria-hidden="true" />No document found</div>'
Set-Content 'src\webparts\spectra\components\SPECTRA.tsx' -Value $content
Write-Host 'Updated no document found with icon'

# DataTable.tsx - Replace sorting icons
$content = Get-Content 'src\webparts\spectra\components\DataTable\DataTable.tsx' -Raw
$content = $content -replace 'className=\`\$\{isAsc \? "fa-solid fa-sort-up" : isDesc \? "fa-solid fa-sort-down" : "fa-regular fa-sort"\} \$\{styles.sortIcons\} \$\{isActive \? styles.sortIconActive : ""\}`', 'src={require(''../../assets/icons/sorting.svg'')} alt="" style={{ width: ''1em'', height: ''1em'', opacity: isActive ? 1 : 0.5, transform: isAsc ? ''scaleY(-1)'' : isDesc ? ''scaleY(1)'' : ''none'' }} className={styles.sortIcons}'
$content = $content -replace '<i (className=.*sorting.*)\s*aria-hidden="true"\s*/>', '<img $1 aria-hidden="true" />'
Set-Content 'src\webparts\spectra\components\DataTable\DataTable.tsx' -Value $content
Write-Host 'Updated sorting icons'

Write-Host 'All icon replacements complete!'
