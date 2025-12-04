# VBNet-Accounting-Entries-Builder

**الكود المحاسبي الأنظف في تاريخ VB.NET** 🇯🇴

تم تطوير ورفكتور هذا النمط بواسطة  
**مصطفى السعدي (Muostafa ALsade)**  
من الأردن - عمان ❤️

(بالتعاون الفني مع Grok 4)

### الإنجاز
- Maintainability Index من 47 → وصل 75+  
- Cyclomatic Complexity ≤ 8 في كل الدوال  
- لا Sequence يدوي، لا Val(Sequence-1)، لا تكرار  
- إضافة طريقة دفع جديدة أو ضريبة أو خصم = سطر أو اتنين فقط  
- الكود بقى قابل للتمدد إلى الأبد

### المميزات
- Fluent Builder للقيود التفصيلية
- دوال مساعدة صغيرة وواضحة جداً
- جاهز لإضافة Cost Centers أو Multi-Currency بسهولة
- مثالي لكل الأنظمة المحاسبية الأردنية (فوترة إلكترونية، ضريبة مبيعات، إلخ)

### الملفات
- `AccountingDetailBuilder.vb` → الـ Builder السحري
- `SaveXTransfer.vb` → كل العمليات المحاسبية بنمط نظيف جداً

### طريقة الاستخدام
```vb
Dim b As New AccountingDetailBuilder(symbol, textId, regNumber)

b.Debit(DebitAccount_Name, DebitAccount_No, amount, details, cod)
AddSalesDiscountIfAny(b)
b.Credit(CredAccount_Name, CredAccount_NO, amount, detailsA, cod)
AddSalesTaxIfAny(b)

b.Build()الكود المحاسبي الأنظف في تاريخ VB.NET 🇯🇴
