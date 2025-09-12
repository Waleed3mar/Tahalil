function prepareAndPrint() {
    // 1. قراءة التاريخ من حقل الإدخال
    const reportDateInput = document.getElementById('reportDate').value;

    // التحقق من أن المستخدم أدخل التاريخ
    if (!reportDateInput) {
        alert('يرجى إدخال تاريخ انتهاء البلاغ أولاً.');
        return; // إيقاف العملية إذا لم يتم إدخال تاريخ
    }

    // تحويل التاريخ للصيغة المطلوبة (يوم / شهر / سنة)
    const date = new Date(reportDateInput);
    const formattedDate = `${date.getDate()} / ${date.getMonth() + 1} / ${date.getFullYear()}`;

    // 2. إنشاء العنوان الكامل حسب طلبك
    const fullTitle = `مرسل لسيادتكم البلاغ الأسبوعي للاصابة بالبهارسيا والفاشيولا المنتهي يوم السبت الموافق ${formattedDate} لادارة الضبعة ومستشفى الضبعة`;

    // 3. قراءة قيم العينات من حقول الإدخال
    const values = {
        B8: document.getElementById('B8').value || 0,
        C8: document.getElementById('C8').value || 0,
        D8: document.getElementById('D8').value || 0,
        E8: document.getElementById('E8').value || 0,
        F8: document.getElementById('F8').value || 0,
        G8: document.getElementById('G8').value || 0,
        H8: document.getElementById('H8').value || 0,
        I8: document.getElementById('I8').value || 0,
    };

    // 4. تحديث جدول التقرير المخفي بالقيم الجديدة
    for (const key in values) {
        document.getElementById(`report-${key}`).innerText = values[key];
    }
    
    // 5. وضع العنوان المخصص في التقرير
    document.querySelector('.report-title').innerText = fullTitle;

    // 6. فتح نافذة الطباعة الخاصة بالمتصفح
    window.print();
}
