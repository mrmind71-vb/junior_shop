INSERT INTO FILE7_20(DOC_NO,item, QUANT, MODEL)
SELECT     '000000' AS Expr1, item, quant , model
FROM         f_bal 



enable  TRIGGER PURCHASE_INSERT ON dbo.file7_20