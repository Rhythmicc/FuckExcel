#!/usr/bin/env python
# -*- coding: utf-8 -*-
name = 'FuckExcel'


def getFuckExcel(excel_path, with_numba=False):
    if with_numba:
        from .FuckExcel_numba import FuckExcel
    else:
        from .FuckExcel import FuckExcel
    return FuckExcel(excel_path)
