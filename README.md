# FuckExcel

- Easier to operate excel.

## Install

```shell
python3 setup.py install
```

## DEMO
### demo - 1

```python
from FuckExcel import getFuckExcel

fuck_excel = getFuckExcel('./A.xlsx', with_numba=True) # default with_numba is False
fuck_excel[5:10, 5:10] = 'init' # or ['init', 'init', 'init', 'init', 'init']
fuck_excel.save()
```

- Demo will create `A.xlsx` and set init value.

![demo](demo.png)

### demo - 2

```python
from FuckExcel import getFuckExcel

fuck_excel = getFuckExcel('./A.xlsx', with_numba=True) # default with_numba is False
fuck_excel[5:, 1] = [1, 2, 3, 4, 5]  # set [5][1]~[10][1] = [1, 2, 3, 4, 5]
fuck_excel.save()
```

