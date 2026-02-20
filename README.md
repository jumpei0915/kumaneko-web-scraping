# kumaneko-web-scraping
This is to manage the tennis court reservation data 

## Read reservation Excel

```bash
python scripts/read_reservation_excel.py
```

Options:
- `--path`: Excel file or directory (default: `/home/jumpei_private/workspace/熊猫カンパニー/reservation`)
- `--sheet`: sheet name or index
- `--rows`: number of preview rows
- `--output`: save as `.csv` or `.json`
