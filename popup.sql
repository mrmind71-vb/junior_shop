SELECT TOP (1000) v.[ORDER_NO]
      ,v.[ORDER_NO_SUP]
      ,[DATE]
      ,[NAME]
      ,[PHONE]
      ,v.[TYPE]
	  ,STAGES_CODES.DESCA
      ,[TOTAL_QUANT]
      ,[TOTAL_ITEM]
      ,[DISCOUNT]
      ,[CLOSED]
      ,[CANCELED]
      ,v.[INV_NO]
	  ,CASE WHEN CLOSED = 1 THEN '„€·ﬁ…' WHEN CANCELED = 1 THEN '„·€«…' ELSE '„› ÊÕ…' END
	  ,s.DATE_PICK
  FROM [JUNIOR_SHOP].[dbo].[vw_online_orders]  v
	INNER JOIN STAGES_CODES ON v.stage = STAGES_CODES.code
	LEFT JOIN FILE6_90S s ON v.ORDER_NO = s.ORDER_NO AND v.ORDER_NO_SUP = s.ORDER_NO_SUP
	WHERE v.[TYPE] IN(0,1,2,3,11,12) 
	order by date desc


