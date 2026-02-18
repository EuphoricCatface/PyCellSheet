---
layout: default
section: manual
parent: ../
title: Overview
---

# Overview

## What is PyCellSheet?

PyCellSheet is a fork of pyspread that computes Python expressions in its cells.
It is written in [Python](https://www.python.org/).

At least basic Python knowledge is helpful when using PyCellSheet effectively.

Like pyspread, PyCellSheet provides a three-dimensional grid (rows, columns, and
tables/sheets). A major difference is PyCellSheet's copy-priority semantics:
referencing another cell returns a deep-copied value by default, which makes
cell behavior closer to conventional spreadsheets for mutable objects.

PyCellSheet also provides spreadsheet-style helper functions and reference
syntax through its expression/reference parser pipeline, while still allowing
direct Python code in cells.
