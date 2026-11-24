[English](cleaning_report.md)

<div dir="rtl">

# تقرير التنظيف - Greenfield Academy، الفصل الأول 2026/27

أُنشئ بواسطة `scripts/02_clean_and_compile.py`. كل صف وراء هذه الأرقام موجود في جدول `cleaning_log`
ضمن `clean_data/school.db`، وفي الأوراق *Corrections* و *Omissions* و *Needs Review* ضمن
`clean_data/school_clean.xlsx`.

## 1. ما قُرئ وما أُنجز

| ملف المصدر | الصيغة | السجلات المقروءة | السجلات المحتفظ بها | standardised | corrected | omitted | missing | flagged | needs_review |
|:---|:---|---:|---:|---:|---:|---:|---:|---:|---:|
| 01_student_roster.xlsx | قوائم طلاب Excel | 244 | 240 | 74 | 14 | 6 | 6 | 0 | 22 |
| 02_enrollment_form_responses.csv | استمارة عبر الإنترنت (CSV) | 241 | 224 | 25 | 60 | 32 | 31 | 0 | 59 |
| 03_gradebook_term1.xlsx | دفتر درجات Excel | 476 | 458 | 16 | 52 | 16 | 6 | 0 | 32 |
| 04_science_department_results.docx | جداول Word | 237 | 222 | 11 | 21 | 15 | 4 | 0 | 21 |
| 05_history_gpa_report.pdf | جدول PDF | 241 | 229 | 14 | 6 | 12 | 2 | 0 | 17 |
| 06_gate_attendance.log | سجل نظام | 9519 | 9308 | 11 | 21 | 71 | 0 | 47 | 83 |

- **standardised (موحَّد)** - الصيغة فقط (المسافات، حالة الأحرف، ترميز التاريخ/المعرّف/الكسر العشري)؛ والقيمة تعني الشيء نفسه.
- **corrected (مصحَّح)** - أُصلحت القيمة أو حُوِّلت: تغيير المقياس، أو خطأ مطبعي، أو مطابقة مع مصدر آخر.
- **omitted (مستبعَد)** - تُرك صف أو قيمة عمدًا: تكرارات، أو إدخالات تجريبية، أو نص غير قابل للاستخدام، أو تمريرات لا يمكن نسبتها إلى أحد.
- **missing (مفقود)** - بيانات كانت المدرسة تتوقعها لكن المصدر لا يحتويها.
- **flagged (مُعلَّم)** - لم يُغيَّر شيء، لكن المُدخل غير معتاد وأُحيل إلى شخص (أحداث سجل البوابة).
- **needs_review (يحتاج إلى مراجعة)** - مدخلات (من أي إجراء) يجب أن يؤكدها شخص؛ انظر القسم 4.

النتيجة: **240 طالبًا**، و**909 من 960 درجة مادة** متاحة على مقياس واحد
0-100، و**4,800 يوم حضور للطلاب** من 9,519 سطرًا في السجل.

## 2. توحيد المقاييس في نسبة مئوية واحدة

| المصدر | المقياس الموجود | القاعدة المطبَّقة |
|---|---|---|
| دفتر الدرجات (الرياضيات، اللغة الإنجليزية) | 0-100، وشعبة واحدة مُدخلة بمقياس 0-1 | الكسر x 100 عندما تؤكده الأعمال + الامتحان |
| العلوم (Word) | 0-1؛ جدول الصف 11 بترويسة "/20"؛ نسب مئوية شاردة | يُقرأ المقياس من ترويسة كل جدول؛ ‎/20 x 5؛ القيم 10-100 في عمود 0-1 تُقرأ نسبة مئوية |
| التاريخ (PDF) | GPA من 0 إلى 4؛ نسب مئوية شاردة | Percent = 50 + 12.5 x GPA (السياسة المطبوعة في التقرير) |
| استمارة التسجيل | نص حر: 88، 88%، 0.88، 3.4، 3.4/4، B+ | ‎<=1 -> x100؛ ‏1-4 -> سياسة GPA (مُعلَّمة كافتراض)؛ الحروف مُعلَّمة |

القيم التي لا تناسب أي مقياس (مثل 8.5 في عمود 0-1، أو GPA بقيمة 4.8) **لا تُخمَّن**: تُحجب الدرجة وتُحال إلى المراجعة.

## 3. حقول النص الحر الموحَّدة بالجملة

| المصدر | الحقل | القيم | عدد الصيغ المختلفة | أمثلة |
|:---|:---|---:|---:|:---|
| 02_enrollment_form_responses.csv | grade | 224 | 5 | `grade N` \| `N` \| `Nth` \| `gN` \| `word` |
| 02_enrollment_form_responses.csv | phone | 224 | 7 | `+N NNN NNN NNNN` \| `(NNN) NNN-NNNN` \| `NNN.NNN.NNNN` \| `NNN-NNN-NNNN` \| `NNNNNNNNNN` |
| 02_enrollment_form_responses.csv | transport | 224 | 12 | `Bus` \| `Car` \| `Walk` \| `Bicycle` \| `bus` |

## 4. مدخلات تحتاج إلى شخص - 234 في المجموع

| الأولوية | الفئة | المدخلات | الطريقة |
|:---|:---|---:|:---|
| عالية | text_absent | 8 | كان الطالب غائبًا عن التقييم. حدِّد موعد امتحان تعويضي، ثم أدخل الدرجة. |
| عالية | unknown_id | 8 | استُخدم عند البوابة معرّف بطاقة غير موجود في السجل. على مكتب الأمن تحديد حامل البطاقة. |
| عالية | odd_hour_swipe | 6 | تسجيل دخول بين 20:00 و 05:00. على مكتب الأمن مراجعة تسجيلات الكاميرات لتلك البوابة والوقت. |
| عالية | dob_illogical | 2 | العمر مستحيل بالنسبة للصف. راجع شهادة الميلاد؛ السنة هي موضع الخطأ المرجَّح. |
| عالية | dob_missing | 2 | استخرج نسخة شهادة الميلاد من ملف القبول وأدخل التاريخ في السجل. |
| عالية | duplicate_conflict | 2 | صفّان بدرجتين مختلفتين لطالب واحد. حُجبت الدرجة؛ على المعلّم تحديد الصف الصحيح. |
| عالية | duplicate_id_conflict | 2 | طالبان يتشاركان معرّفًا واحدًا في السجل. أُعيد إسناد المعرّف بمطابقة الاسم مع استمارة التسجيل؛ على المسجّل التأكيد وتصحيح السجل. |
| متوسطة | no_form_response | 16 | لا توجد استمارة تسجيل في الملف، وبالتالي لا بيانات تواصل لولي الأمر. على معلّم الفصل إرسال الاستمارة مجددًا. |
| متوسطة | text_not_available | 16 | لا درجة مسجّلة (N/A، قيد الانتظار، معفى، غير مكتمل). على معلّم المادة تأكيد الحالة النهائية. |
| متوسطة | early_leave | 12 | غادر قبل 12:00. طابِق مع أذونات الانصراف المبكر الموقّعة. |
| متوسطة | row_missing | 12 | الطالب مدرج في السجل لكن لا صف له في قائمة الدرجات هذه. على معلّم المادة تقديم الدرجة أو السبب. |
| متوسطة | letter_grade | 9 | تقدير بالحروف حيث يلزم رقم؛ ولا يوجد تحويل دقيق. اطلب الدرجة الرقمية من المعلّم/المدرسة السابقة. |
| متوسطة | absent_but_swiped | 8 | سجّله النظام غائبًا (ABSENT) لكن البطاقة مُرِّرت. احتُسب حاضرًا؛ على معلّم الفصل التأكيد. |
| متوسطة | total_mismatch | 6 | المجموع لا يساوي الأعمال + الامتحان. استُخدم مجموع المكوّنات؛ على المعلّم مراجعة كشف الدرجات الورقي. |
| متوسطة | blank | 5 | عدد أرقام الهاتف غير صحيح. أكِّده عبر البريد الإلكتروني أو من خلال الطالب. |
| متوسطة | illogical_value | 5 | القيمة خارج النطاق الممكن. راجع الاستمارة الورقية الأصلية أو كشف الدرجات. |
| متوسطة | weekend_swipe | 5 | تمرير في يوم غير دراسي. أكِّده مقابل تقويم الأنشطة. |
| متوسطة | component_out_of_range | 4 | أحد المكوّنات يتجاوز درجته القصوى. أُبقي المجموع كما سُجِّل؛ على المعلّم إعادة إدخال المكوّن. |
| متوسطة | dob_conflict | 3 | السجل واستمارة التسجيل غير متفقين. أُبقيت قيمة السجل؛ أكِّدها بشهادة الميلاد. |
| متوسطة | email_invalid | 3 | العنوان لا يمكنه استقبال البريد. اتصل بولي الأمر لتأكيد العنوان. |
| متوسطة | gpa_over_4 | 3 | قيمة GPA أعلى من 4.0 مستحيلة وفق سياسة المدرسة. حُجبت الدرجة؛ على معلّم التاريخ إعادة إدخالها. |
| متوسطة | misplaced_decimal | 3 | القيمة لا تناسب مقياس 0-1 ولا مقياس 0-100؛ والأرجح أن الفاصلة العشرية في غير موضعها. على المعلّم التأكيد (القيمة المقترحة في عمود التفاصيل). |
| متوسطة | no_record | 3 | لا تمرير ولا إدخال ABSENT في ذلك اليوم. احتُسب غائبًا؛ أكِّد ذلك من سجل الفصل. |
| متوسطة | phone_invalid | 3 | عدد أرقام الهاتف غير صحيح. أكِّده عبر البريد الإلكتروني أو من خلال الطالب. |
| متوسطة | unknown_id | 3 | المعرّف غير موجود في السجل. تحقَّق من وجود تسجيل متأخر أو معرّف مكتوب خطأً قبل الاستبعاد. |
| متوسطة | wrong_id | 3 | حملت الاستمارة معرّف طالب آخر. أُعيد الربط بمطابقة تامة للاسم؛ أكِّد ذلك مع جهة اتصال ولي الأمر. |
| متوسطة | out_of_range | 2 | الدرجة أعلى من الحد الأقصى للمقياس المعلن. على المعلّم تأكيد القيمة والمقياس المستخدم. |
| منخفضة | missing_checkout | 21 | لا تسجيل خروج. تأكَّد من أن الطالب غادر عبر بوابة مراقَبة. |
| منخفضة | missing_checkin | 14 | سجّل خروجًا دون تسجيل دخول. على معلّم الفصل تأكيد وقت الوصول من سجل الفصل. |
| منخفضة | scale_gpa | 12 | قيمة بين 1 و 4 قُرئت كمعدل GPA بمقياس 4.0 وحُوِّلت وفق سياسة المدرسة. أكِّدها مقابل كشف درجات المدرسة السابقة. |
| منخفضة | dob_missing | 6 | أُخذ التاريخ من استمارة التسجيل التي ملأها ولي الأمر. أكِّده مقابل شهادة الميلاد في التدقيق القادم للملفات. |
| منخفضة | scale_percent_in_gpa_column | 6 | كُتبت نسبة مئوية في عمود GPA؛ وحُوِّلت وفق سياسة المدرسة. على المعلّم التأكيد. |
| منخفضة | gender_missing | 4 | الحقل فارغ في السجل. أكمِله من استمارة القبول. |
| منخفضة | id_empty | 4 | لا يمكن نسبة تمرير بمعرّف فارغ إلى أحد. افحص قارئ البوابة/البطاقة. |
| منخفضة | dob_illogical | 3 | أُخذ التاريخ من استمارة التسجيل التي ملأها ولي الأمر. أكِّده مقابل شهادة الميلاد في التدقيق القادم للملفات. |
| منخفضة | grade_mismatch | 3 | خانة الصف لا تتفق مع الورقة المدرج فيها الطالب. أُبقيت قيمة الورقة؛ على المسجّل تأكيد التسكين. |
| منخفضة | name_mismatch | 3 | التهجئة في الاستمارة تختلف عن السجل. أُبقيت تهجئة السجل؛ اسأل ولي الأمر عن التهجئة الصحيحة. |
| منخفضة | id_missing_name_typo | 2 | الصف بلا معرّف والاسم لا يطابق تمامًا؛ رُبط بأقرب اسم ضمن الصف نفسه. أكِّد الربط. |
| منخفضة | impossible_date | 2 | طابع زمني خارج فترة التصدير (عطل في الساعة). على قسم تقنية المعلومات فحص ساعة جهاز التحكم في البوابة. |

القائمة الكاملة، بصف واحد لكل مُدخل مع القيمة الخام والطالب، موجودة في ورقة *Needs Review*
(مرتّبة من الأولوية العالية إلى المنخفضة). منها 83 مُدخلًا مصدرها سجل البوابة، وهي مدرجة أيضًا مع
التاريخ والوقت في ورقة *Log Flags*.

## 5. كل شيء حسب المصدر والإجراء والفئة

قيم الإجراء والفئة هنا هي الرموز كما تظهر في جدول `cleaning_log` (انظر معاني الإجراءات في القسم 1).

| المصدر | الإجراء | الفئة | المدخلات |
|:-----------------------------------|:-------------|:---------------------------------|----------:|
| 01_student_roster.xlsx             | corrected    | dob_illogical                    |         3 |
| 01_student_roster.xlsx             | corrected    | dob_missing                      |         6 |
| 01_student_roster.xlsx             | corrected    | duplicate_id_conflict            |         2 |
| 01_student_roster.xlsx             | corrected    | grade_mismatch                   |         3 |
| 01_student_roster.xlsx             | missing      | dob_missing                      |         2 |
| 01_student_roster.xlsx             | missing      | gender_missing                   |         4 |
| 01_student_roster.xlsx             | omitted      | dob_illogical                    |         2 |
| 01_student_roster.xlsx             | omitted      | duplicate_row                    |         4 |
| 01_student_roster.xlsx             | standardised | dob_text_format                  |        20 |
| 01_student_roster.xlsx             | standardised | gender_variant                   |        18 |
| 01_student_roster.xlsx             | standardised | id_format                        |        14 |
| 01_student_roster.xlsx             | standardised | name_case_space                  |        16 |
| 01_student_roster.xlsx             | standardised | section_format                   |         6 |
| 02_enrollment_form_responses.csv   | corrected    | dob_conflict                     |         3 |
| 02_enrollment_form_responses.csv   | corrected    | email_typo                       |         3 |
| 02_enrollment_form_responses.csv   | corrected    | name_mismatch                    |         3 |
| 02_enrollment_form_responses.csv   | corrected    | number_as_word                   |        16 |
| 02_enrollment_form_responses.csv   | corrected    | number_with_text                 |         3 |
| 02_enrollment_form_responses.csv   | corrected    | scale_fraction                   |        12 |
| 02_enrollment_form_responses.csv   | corrected    | scale_gpa                        |        17 |
| 02_enrollment_form_responses.csv   | corrected    | wrong_id                         |         3 |
| 02_enrollment_form_responses.csv   | missing      | blank                            |        15 |
| 02_enrollment_form_responses.csv   | missing      | no_form_response                 |        16 |
| 02_enrollment_form_responses.csv   | omitted      | duplicate_submission             |        12 |
| 02_enrollment_form_responses.csv   | omitted      | email_invalid                    |         3 |
| 02_enrollment_form_responses.csv   | omitted      | illogical_value                  |         4 |
| 02_enrollment_form_responses.csv   | omitted      | letter_grade                     |         5 |
| 02_enrollment_form_responses.csv   | omitted      | phone_invalid                    |         3 |
| 02_enrollment_form_responses.csv   | omitted      | test_submission                  |         3 |
| 02_enrollment_form_responses.csv   | omitted      | unknown_id                       |         2 |
| 02_enrollment_form_responses.csv   | standardised | email_case_space                 |        10 |
| 02_enrollment_form_responses.csv   | standardised | scale_percent_sign               |        15 |
| 03_gradebook_term1.xlsx            | corrected    | component_out_of_range           |         4 |
| 03_gradebook_term1.xlsx            | corrected    | id_missing                       |         8 |
| 03_gradebook_term1.xlsx            | corrected    | illogical_total                  |         8 |
| 03_gradebook_term1.xlsx            | corrected    | scale_fraction                   |        20 |
| 03_gradebook_term1.xlsx            | corrected    | total_blank                      |         6 |
| 03_gradebook_term1.xlsx            | corrected    | total_mismatch                   |         6 |
| 03_gradebook_term1.xlsx            | missing      | row_missing                      |         6 |
| 03_gradebook_term1.xlsx            | omitted      | duplicate_conflict               |         2 |
| 03_gradebook_term1.xlsx            | omitted      | text_absent                      |         8 |
| 03_gradebook_term1.xlsx            | omitted      | text_not_available               |         6 |
| 03_gradebook_term1.xlsx            | standardised | number_as_text                   |        16 |
| 04_science_department_results.docx | corrected    | id_missing                       |         5 |
| 04_science_department_results.docx | corrected    | id_missing_name_typo             |         2 |
| 04_science_department_results.docx | corrected    | scale_percent_in_fraction_column |        14 |
| 04_science_department_results.docx | missing      | row_missing                      |         4 |
| 04_science_department_results.docx | omitted      | letter_grade                     |         4 |
| 04_science_department_results.docx | omitted      | misplaced_decimal                |         3 |
| 04_science_department_results.docx | omitted      | out_of_range                     |         2 |
| 04_science_department_results.docx | omitted      | text_not_available               |         5 |
| 04_science_department_results.docx | omitted      | unknown_id                       |         1 |
| 04_science_department_results.docx | standardised | comma_decimal                    |         6 |
| 04_science_department_results.docx | standardised | id_letter_O                      |         5 |
| 05_history_gpa_report.pdf          | corrected    | scale_percent_in_gpa_column      |         6 |
| 05_history_gpa_report.pdf          | missing      | row_missing                      |         2 |
| 05_history_gpa_report.pdf          | omitted      | duplicate_row                    |         3 |
| 05_history_gpa_report.pdf          | omitted      | gpa_over_4                       |         3 |
| 05_history_gpa_report.pdf          | omitted      | illogical_value                  |         1 |
| 05_history_gpa_report.pdf          | omitted      | text_not_available               |         5 |
| 05_history_gpa_report.pdf          | standardised | comma_decimal                    |         5 |
| 05_history_gpa_report.pdf          | standardised | id_format                        |         5 |
| 05_history_gpa_report.pdf          | standardised | number_as_text                   |         4 |
| 06_gate_attendance.log             | corrected    | absent_but_swiped                |         8 |
| 06_gate_attendance.log             | corrected    | late_not_marked                  |        10 |
| 06_gate_attendance.log             | corrected    | no_record                        |         3 |
| 06_gate_attendance.log             | flagged      | early_leave                      |        12 |
| 06_gate_attendance.log             | flagged      | missing_checkin                  |        14 |
| 06_gate_attendance.log             | flagged      | missing_checkout                 |        21 |
| 06_gate_attendance.log             | omitted      | duplicate_swipe                  |        40 |
| 06_gate_attendance.log             | omitted      | id_empty                         |         4 |
| 06_gate_attendance.log             | omitted      | impossible_date                  |         2 |
| 06_gate_attendance.log             | omitted      | odd_hour_swipe                   |         6 |
| 06_gate_attendance.log             | omitted      | truncated_line                   |         6 |
| 06_gate_attendance.log             | omitted      | unknown_id                       |         8 |
| 06_gate_attendance.log             | omitted      | weekend_swipe                    |         5 |
| 06_gate_attendance.log             | standardised | id_format                        |        10 |
| 06_gate_attendance.log             | standardised | timestamp_format                 |         1 |

</div>
