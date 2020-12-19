---
layout: post
title: 通过宏禁止Excel文件的保存
slug: do-not-save-my-excel
date: 2019-06-21 21:01:58
status: publish
author: <Huasing>
categories: 
  - 默认分类
tags: 
  - Excel
  - VBA
excerpt: 在工作中，一些Excel表格只用于查询和临时修改，实际并不需要修改保存
---

在工作中，一些Excel表格只用于查询和临时修改，实际并不需要修改保存。为避免一时手残给保存了，就需要做出限制。
  
VBA**设计模式**下，在ThisWorkbook中插入VBA代码：
```
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Me.Saved = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Cancel = True
End Sub

```
保存后返回Excel即可。
