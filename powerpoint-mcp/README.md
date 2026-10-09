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
