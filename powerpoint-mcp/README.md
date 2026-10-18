# PowerPoint Live Automation MCP Server 🎬

خادم **MCP (Model Context Protocol)** يُمكّن أدوات الذكاء الاصطناعي (مثل **Antigravity**, **Codex**, **Claude**, **Cursor**) من التحكم المباشر والحي في برنامج **Microsoft PowerPoint** على نظام Windows.

الخادم يعتمد على تقنية **Windows COM Automation (`pywin32`)**، مما يجعل نافذة PowerPoint تفتح وتتحرك مباشرة على شاشتك، وترى النصوص، الأشكال، والبطاقات تُبنى وتتحرك خطوة بخطوة في الوقت الفعلي!

---

## 🚀 المميزات الرئيسية:
1. **عرض حي وتفاعلي (Live View)**: لا يتم إنشاء ملفات صامتة في الخلفية فقط، بل تراقب نافذة PowerPoint وهي ترسم الشرائح مباشرة على سطح مكتبك.
2. **تصاميم عصرية (Modern UI Cards & Gradients)**: دعم الألوان الحديثة (Dark mode, Slate, Blue)، البطاقات المربعة والدائرية، التدرجات اللونية، وتنسيقات الـ 16:9 الاحترافية.
3. **حركات ومؤثرات حية (Animations & Transitions)**: إضافة مؤثرات الدخول (Fade, Zoom, Bounce, Fly-in) والانتقالات بين الشرائح.
4. **تحكم كامل في العرض (Live Slideshow)**: بدء العرض التقديمي بحجم الشاشة الكاملة، والتنقل بين الشرائح (Next, Previous, Exit) عن بُعد عبر الأوامر.
5. **جداول ونقاط وملاحظات المتحدث**: إنشاء جداول منسقة وقوائم نقطية وملاحظات سرية للمتحدث (Speaker Notes).

---

## 🛠️ كيفية تفعيل الـ MCP Server

### 1. في Antigravity / Claude Desktop
أضف الإعدادات التالية إلى ملف التكوين الخاص بـ MCP (مثلاً `claude_desktop_config.json` أو إعدادات Antigravity):

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\powerpoint-mcp",
        "python",
        "server.py"
      ]
    }
  }
}
```

أو باستخدام Python المباشر:
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\powerpoint-mcp\\server.py"
      ]
    }
  }
}
```

---

## 🧰 الأدوات المتاحة (Available Tools):

| الأداة | الوظيفة |
| :--- | :--- |
| `ppt_status` | فحص حالة PowerPoint والعرض التقديمي المفتوح وعدد الشرائح |
| `ppt_new_presentation` | إنشاء عرض تقديمي جديد بأبعاد 16:9 وفتحه على الشاشة |
| `ppt_open` | فتح ملف PowerPoint موجود مسبقاً (.pptx) |
| `ppt_save` | حفظ العرض التقديمي كملف `.pptx` أو تصديره إلى `.pdf` |
| `ppt_close` | إغلاق العرض التقديمي الحالي |
| `ppt_add_slide` | إضافة شريحة جديدة والتوجه إليها مباشرة لترى التغيير |
| `ppt_goto_slide` | التنقل إلى شريحة معينة على الشاشة |
| `ppt_set_background` | تعيين لون خلفية الشريحة (ألوان أحادية أو Gradient متدرج) |
| `ppt_add_textbox` | إضافة نص منسق (حجم الخط، اللون، المحاذاة، العريض/المائل) |
| `ppt_add_card` | إضافة بطاقة تصميم عصرية (UI Card) مع عنوان ووصف وشريط تمييز |
| `ppt_add_shape` | إضافة أشكال هندسية (مستطيل، دائرة، أسهم، نجوم، سداسي...) |
| `ppt_add_bullet_list` | إضافة قائمة نقطية منسقة |
| `ppt_add_table` | إضافة جدول بيانات مع ألوان مخصصة للرأس والصفوف |
| `ppt_add_image` | إدراج صورة من القرص الصلب |
| `ppt_add_animation` | إضافة أنيميشن للشكل (Fade, Zoom, Fly, Bounce...) |
| `ppt_set_transition` | تعيين حركة الانتقال بين الشرائح (Fade, Push, Wipe...) |
| `ppt_set_speaker_notes`| إضافة ملاحظات المتحدث للشريحة |
| `ppt_add_metric_card` | إضافة بطاقة إحصائيات وKPIs برقم عريض ومؤشر الأداء |
| `ppt_add_timeline` | إضافة مسار زمني ومراحل (Roadmap / Process Flow) مع ترقيم ودوائر |
| `ppt_add_code_block` | نافذة محرّر أكواد للمطورين (أزرار Mac، خط Consolas، وثيم داكن) |
| `ppt_add_bar_chart` | إضافة رسم بياني شريطي (Vector Bar Chart) بالنسب والألوان |
| `ppt_add_badge` | إضافة وسام أو تاج تصنيف (Badge / Pill) للأقسام والحالات |
| `ppt_export_slide_image`| تصدير أي شريحة كصورة عالية الدقة PNG / JPG بدقة 1080p |
| `ppt_duplicate_slide` | نسخ ومضاعفة شريحة موجودة بكل تفاصيلها |
| `ppt_move_slide` | نقل وتغيير ترتيب الشرائح داخل العرض |
| `ppt_read_slide_content`| قراءة واستخراج جميع النصوص والأشكال والجداول من الشريحة |
| `ppt_search_and_replace_text`| البحث والاستبدال الذكي للنصوص والقوالب (`{{VARIABLE}}`) عبر جميع الشرائح |
| `ppt_add_quote_card` | بطاقة اقتباس أو شهادة عميل (Testimonial Quote) مع اسم القائل والصورة |
| `ppt_add_pros_cons` | مقارنة جانبية ببطاقتين للإيجابيات والسلبيات (Pros & Cons) |
| `ppt_add_pricing_table`| جدول خطط الأسعار والاشتراكات (Pricing Table) مع تمييز الباقة الأكثر طلباً |
| `ppt_add_donut_chart` | رسم دائري لمؤشر التقدم والنسب المئوية (Donut Progress Chart) |
| `ppt_add_callout_box` | صندوق تنبيه / نصيحة / تحذير (Alert & Callout Box) مع شريط جانبي ملون وأيقونة تصنيف |
| `ppt_add_process_steps`| خط سير عمليات ومراحل مرقمة (Numbered Workflow Process) مع أسهم ربط حديثة |
| `ppt_add_feature_grid` | شبكة مصفوفة متعددة الأعمدة لبطاقات المميزات والخدمات (Feature Grid) |
| `ppt_add_team_card` | بطاقة تعريف بأعضاء الفريق والمتحدثين (Team Profile Card) بالأفاتار والمهارات |
| `ppt_add_agenda` | جدول أعمال الاجتماع أو العرض (Meeting Agenda / TOC) مع تمييز الفقرة الحالية |
| `ppt_add_watermark` | علامة مائية قطرية شبه شفافة (Watermark: DRAFT / CONFIDENTIAL) خلف الشرائح |
| `ppt_set_shape_zorder` | التحكم في ترتيب طبقات العناصر (Bring to Front / Send to Back) |
| `ppt_add_comparison_matrix`| جدول مصفوفة المقارنة المتقدم (Feature Comparison Matrix) مع تمييز الباقة المختارة |
| `ppt_add_swot_matrix` | مصفوفة التحليل الاستراتيجي الرباعي (SWOT Matrix 2x2: S, W, O, T) بتنسيق ملون أنيق |
| `ppt_add_funnel` | قمع التسويق والتحويل (Sales & Conversion Funnel) مع النسب التنازلية |
| `ppt_add_before_after` | مقارنة التحول والنتائج قبل وبعد (Before vs After) مع سهم التحول |
| `ppt_add_stat_hero` | بطاقة الرقم القياسي البارز (Hero KPI Stat) مع نسبة النمو وشارة الاتجاه |
| `ppt_add_faq_accordion`| قائمة الأسئلة الشائعة والأجوبة (FAQ & Q&A Accordion) ببطاقات منظمة |
| `ppt_add_checklist` | قائمة مهام وتحقق مع شارات الحالة (Checklist: Done, In Progress, Pending) |
| `ppt_add_review_card` | بطاقة تقييم وشهادة 5 نجوم (5-Star Rating & Trust Card) وشارة المنصة |
| `ppt_style_shape` | تنسيق وتخصيص متقدم للأشكال (الشفافية Opacity، الظلال Shadow، الحدود واللون) |
| `ppt_split_slide_to_progressive_deck` | تقسيم وتفجير شريحة واحدة متعددة العناصر إلى عدة شرائح متتالية يبرز فيها عنصر واحد في كل خطوة مع تبهيت الباقي |
| `ppt_create_progressive_deck` | إنشاء عرض تقدمي كامل (Storyline Deck) مع شريط تتبع علوي وبطاقات تتنقل بين المكتمل والنشط والمستقبلي |
| `ppt_animate_sequence` | تشغيل حركات دخول متسلسلة للأشكال والبطاقات بنقرة واحدة أو بتأخير زمني (On-Click / Delay) |
| `ppt_animate_bullet_points` | إظهار عناصر القوائم النقطية فقرة بفقرة (Paragraph Reveal) تفاعلياً عند كل نقرة |
| `ppt_animate_emphasis` | حركات تأكيد وجذب الانتباه (Emphasis: Pulse, Grow/Shrink, Spin, Teeter) لأرقام وبطاقات معينة أثناء الشرح |
| `ppt_animate_exit` | حركات خروج وتلاشي (Exit Animations: Fade Out, Disappear, Fly Out) لإخفاء العناصر السابقة |
| `ppt_clear_animations` | مسح وتنظيف كل الحركات والمؤثرات من الشريحة لإعادة ضبطها |
| `ppt_add_animated_stepper` | رسم مسار خطوات تفاعلي (Animated Stepper Pipeline) مجهز بحركات دخول متزامنة تلقائياً |
| `ppt_add_venn_diagram` | رسم مخطط فين (Venn Diagram) بدائرتين متداخلتين شفافتين مع شارة التآزر الاستراتيجي (Synergy Badge) |
| `ppt_add_pyramid_hierarchy` | رسم هرم استراتيجي وتسلسل هرمي متعدد المستويات (Strategy / Maslow Pyramid Stack) |
| `ppt_add_matrix_2x2` | مصفوفة الأولويات الرباعية 2x2 (Gartner Magic Quadrant / Impact-Effort Matrix) مع محاور X و Y |
| `ppt_add_roadmap_milestones` | مسار خارطة طريق ربع سنوية (Quarterly Roadmap & Milestones) بنقاط مسار وبطاقات مخرجات |
| `ppt_add_user_persona` | بطاقة شخصية العميل والمستخدم (User Persona Card) بالأفاتار، الأهداف، نقاط الألم، وحزمة الأدوات |
| `ppt_add_roi_card` | بطاقة دراسة الجدوى والعائد على الاستثمار وحساب الوفر المالي (Executive ROI & Cost Savings Case) |
| `ppt_add_dashboard_strip` | شريط لوحة قيادة موحد متعدد المؤشرات الرقمية (KPI Dashboard Strip) وشارات النمو الإيجابي والسلبي |
| `ppt_batch_format_deck` | توحيد وتنسيق الخطوط وألوان العناوين وتذييل الصفحات دفعة واحدة على جميع شرائح العرض بنقرة واحدة |
| `ppt_delete_shape` | حذف شكل معين من الشريحة عبر الاسم أو المعرف (ID) |
| `ppt_clear_slide` | مسح وتنظيف كل الأشكال من الشريحة لإعادة تعيينها |
| `ppt_add_footer` | إضافة تذييل احترافي (Footer) مع رقم الشريحة وعبارة السرية على كل الشرائح |
| `ppt_run_slideshow` | تشغيل العرض التقديمي في وضع ملء الشاشة (Slideshow) |
| `ppt_slideshow_next` | الانتقال للشريحة أو الحركة التالية أثناء العرض |
| `ppt_slideshow_previous`| الرجوع للشريحة السابقة أثناء العرض |
| `ppt_slideshow_exit` | إنهاء وضع ملء الشاشة والرجوع لوضع التحرير |

---

## 🧪 تجربة العرض الحي (Demo):
لتشغيل تجربة حية ومشاهدة PowerPoint وهو يبني الشرائح أمام عينيك:
```bash
uv run python test_demo.py
```
