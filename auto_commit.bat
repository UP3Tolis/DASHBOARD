cd "C:\Users\PLN\OneDrive - PLN\002. Perencanaan UP3 TLI\999. DASHBOARD\DASHBOARD"   # <- ganti dengan path kamu

git add .
git commit -m "Auto update at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 
git push

git add --force "NKO UP3 TLI.xlsx"
git commit -m "Force update excel"
git push
