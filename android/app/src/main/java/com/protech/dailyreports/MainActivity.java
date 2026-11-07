package com.protech.dailyreports;
import android.app.Activity;
import android.os.Bundle;
import android.app.AlertDialog;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.drawable.GradientDrawable;
import android.view.WindowInsets;
import android.graphics.Insets;
import android.text.TextWatcher;
import android.text.Editable;
import android.net.Uri;
import android.view.View;
import android.view.Gravity;
import android.view.ViewGroup;
import android.widget.*;
import android.text.InputType;
import org.json.*;
import java.util.*;
import java.text.SimpleDateFormat;
import java.io.*;

public class MainActivity extends Activity {
    JSONObject state, config, categories;
    JSONArray order, reports;
    LinearLayout body;
    static final int TEAL=Color.rgb(37,75,77), LIGHT=Color.rgb(244,248,247);
    static final int NAVY=Color.rgb(24,62,67), GREEN=Color.rgb(34,133,103);
    static final int WHITE=Color.WHITE, MUTED=Color.rgb(93,118,120), BORDER=Color.rgb(220,234,231);
    int dp(int value){ return (int)(getResources().getDisplayMetrics().density*value+0.5f); }
    GradientDrawable background(int color,int radius){
        GradientDrawable d=new GradientDrawable();
        d.setColor(color);d.setCornerRadius(dp(radius));return d;
    }
    final int importCode=727;

    @Override public void onCreate(Bundle b) {
        super.onCreate(b);
        try { state=new JSONObject(getPreferences(0).getString("state","")); }
        catch(Exception e) { state=new JSONObject(); }
        if(state.optJSONObject("config")==null) setupDefaults();
        config=state.optJSONObject("config");
        categories=config.optJSONObject("categories");
        order=config.optJSONArray("category_order");
        reports=state.optJSONArray("reports");
        if(reports==null){ reports=new JSONArray(); put(state,"reports",reports); }
        home();
    }
    void put(JSONObject obj,String k,Object v){try{obj.put(k,v);}catch(Exception ignored){}}
    void persist(){
        boolean okay=getPreferences(0).edit().putString("state",state.toString()).commit();
        if(!okay) Toast.makeText(this,"تعذر حفظ البيانات",Toast.LENGTH_LONG).show();
    }
    JSONArray items(String... a) { JSONArray ar=new JSONArray(); for(String s:a)ar.put(s); return ar; }
    JSONObject field(String id,String label,String kind) {
        JSONObject r=new JSONObject();put(r,"id",id);put(r,"label",label);put(r,"kind",kind);return r;
    }
    JSONObject cat(String title,String person,JSONArray names,JSONObject... fs) {
        JSONObject row=new JSONObject();
        put(row,"label",title);put(row,"person_label",person);put(row,"total_label","إجمالي الأعمال");put(row,"people",names);
        JSONArray f=new JSONArray();for(JSONObject x:fs) f.put(x);put(row,"fields",f);return row;
    }
    void setupDefaults() {
        state=new JSONObject();
        JSONObject cfg=new JSONObject(),cats=new JSONObject();
        put(cfg,"version",11);
        put(cfg,"report_title","");
        put(cfg,"category_order",new JSONArray());
        put(cfg,"categories",cats);
        put(state,"config",cfg);
        put(state,"reports",new JSONArray());
        persist();
    }
    ArrayList<String> ids() {
        ArrayList<String> list=new ArrayList<>();
        if(order!=null)for(int i=0;i<order.length();i++){
            String id=order.optString(i);if(categories.optJSONObject(id)!=null && !list.contains(id))list.add(id);
        }
        Iterator<String> it=categories.keys();
        while(it.hasNext()){String id=it.next();if(!list.contains(id))list.add(id);}
        return list;
    }
    String title(String id){JSONObject c=categories.optJSONObject(id);return c==null?id:c.optString("label",id);}
    JSONObject catFor(String id){return categories.optJSONObject(id);}
    String stamp(){return new SimpleDateFormat("yyyy-MM-dd HH:mm",Locale.US).format(new Date());}

    void screen(String heading) {
        LinearLayout outer=new LinearLayout(this);outer.setOrientation(1);
        outer.setLayoutDirection(View.LAYOUT_DIRECTION_RTL);outer.setBackgroundColor(LIGHT);
        outer.setPadding(dp(12),dp(8),dp(12),dp(8));
        if(android.os.Build.VERSION.SDK_INT>=30){
            outer.setOnApplyWindowInsetsListener((view,insets)->{
                Insets bars=insets.getInsets(WindowInsets.Type.systemBars());
                outer.setPadding(dp(12)+bars.left,dp(8)+bars.top,dp(12)+bars.right,dp(8)+bars.bottom);
                return insets;
            });
        }
        LinearLayout headingBox=new LinearLayout(this);headingBox.setOrientation(1);
        headingBox.setPadding(dp(14),dp(11),dp(14),dp(12));
        headingBox.setBackground(background(NAVY,14));
        TextView h=new TextView(this);h.setText(heading);h.setTextSize(21);h.setTextColor(WHITE);
        h.setGravity(Gravity.RIGHT);h.setTypeface(null,1);headingBox.addView(h);
        TextView sub=new TextView(this);sub.setText("ProTech • إدارة تقارير العمل • v15.5");
        sub.setTextSize(11);sub.setTextColor(Color.rgb(199,231,227));sub.setGravity(Gravity.RIGHT);
        headingBox.addView(sub);outer.addView(headingBox);
        ScrollView sc=new ScrollView(this);
        sc.setClipToPadding(false);sc.setFillViewport(true);
        body=new LinearLayout(this);body.setOrientation(1);
        body.setLayoutDirection(View.LAYOUT_DIRECTION_RTL);
        body.setPadding(0,dp(12),0,dp(16));
        sc.addView(body);outer.addView(sc,new LinearLayout.LayoutParams(-1,0,1));
        LinearLayout nav=new LinearLayout(this);nav.setOrientation(0);
        nav.setPadding(dp(2),dp(6),dp(2),dp(2));nav.setBackground(background(WHITE,14));
        small(nav,"الرئيسية",()->home());
        small(nav,"التقارير",()->history());
        small(nav,"الإدارة",()->chooseCategory(false));
        small(nav,"الإعدادات",()->settingsScreen());
        outer.addView(nav);
        setContentView(outer);
    }
    void small(LinearLayout group,String label,Runnable cb) {
        Button b=new Button(this);b.setText(label);b.setAllCaps(false);b.setTextSize(12);
        b.setTextColor(TEAL);b.setBackground(background(LIGHT,10));
        b.setMinHeight(dp(50));b.setPadding(dp(1),dp(3),dp(1),dp(3));
        LinearLayout.LayoutParams p=new LinearLayout.LayoutParams(0,dp(54),1);
        p.setMargins(dp(2),dp(2),dp(2),dp(2));group.addView(b,p);
        b.setOnClickListener(v->cb.run());
    }
    void button(String label,Runnable cb) {
        Button b=new Button(this);b.setText(label);b.setAllCaps(false);
        b.setTextSize(15);b.setGravity(Gravity.CENTER_VERTICAL|Gravity.RIGHT);
        b.setTextColor(WHITE);
        int tone=label.contains("حذف") || label.contains("إزالة") ? Color.rgb(154,76,72):
                  label.contains("حفظ") || label.contains("إضافة") ? GREEN : TEAL;
        b.setBackground(background(tone,12));
        b.setMinHeight(dp(54));b.setPadding(dp(16),dp(9),dp(16),dp(9));
        LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(-1,-2);
        lp.setMargins(dp(3),dp(4),dp(3),dp(4));body.addView(b,lp);
        b.setOnClickListener(v->cb.run());
    }
    void label(String value) {
        TextView t=new TextView(this);t.setText(value);t.setTextColor(TEAL);
        t.setTextSize(15);t.setTypeface(null,1);t.setGravity(Gravity.RIGHT);
        t.setPadding(dp(10),dp(14),dp(10),dp(7));body.addView(t);
    }
    void statTile(LinearLayout row,String headline,String value) {
        LinearLayout box=new LinearLayout(this);box.setOrientation(1);
        box.setPadding(dp(14),dp(12),dp(14),dp(12));box.setBackground(background(WHITE,12));
        TextView name=new TextView(this);name.setText(headline);name.setTextSize(12);
        name.setTextColor(MUTED);name.setGravity(Gravity.RIGHT);box.addView(name);
        TextView count=new TextView(this);count.setText(value);count.setTextSize(23);
        count.setTypeface(null,1);count.setTextColor(TEAL);count.setGravity(Gravity.RIGHT);box.addView(count);
        LinearLayout.LayoutParams p=new LinearLayout.LayoutParams(0,dp(94),1);
        p.setMargins(dp(3),dp(3),dp(3),dp(3));row.addView(box,p);
    }
    void settingsScreen() {
        screen("الإعدادات والنسخ الاحتياطي");
        label("التخصيص");
        button("تعيين عنوان التقرير",()->editText("اكتب عنوان التقرير",
            config.optString("report_title",""),v->{put(config,"report_title",v);persist();settingsScreen();}));
        button("إدارة الأقسام",()->sections());
        button("إدارة الموظفين",()->chooseCategory(false));
        button("تعديل الحقول والقوائم",()->chooseCategory(true));
        label("حماية البيانات");
        button("مشاركة نسخة احتياطية JSON",()->shareData());
        button("استيراد إعدادات الكمبيوتر config.json",()->importConfig());
        label("تُحفظ بيانات الجوال محليًا، ولا تتم مزامنتها تلقائيًا مع الكمبيوتر.");
        label("Developer: Abu Omar");
        label("Email: protech.2025.2026@gmail.com");
        label("Phone: +966503276927");
        button("مراسلة المطور",()->{
            try{startActivity(new Intent(Intent.ACTION_SENDTO,Uri.parse("mailto:protech.2025.2026@gmail.com")));}
            catch(Exception ex){Toast.makeText(this,"لا يوجد تطبيق بريد",Toast.LENGTH_SHORT).show();}
        });
        button("الاتصال بالمطور",()->{
            try{startActivity(new Intent(Intent.ACTION_DIAL,Uri.parse("tel:+966503276927")));}
            catch(Exception ex){Toast.makeText(this,"لا يوجد تطبيق اتصال",Toast.LENGTH_SHORT).show();}
        });
    }
    void editText(String question,String initial,java.util.function.Consumer<String> cb) {
        EditText entry=new EditText(this);entry.setSingleLine(true);
        entry.setText(initial);entry.setHint(question);entry.setGravity(Gravity.RIGHT);
        new AlertDialog.Builder(this).setTitle(question).setView(entry).setNegativeButton("إلغاء",null)
        .setPositiveButton("حفظ",(dlg,which)->{
            String s=entry.getText().toString().trim();
            if(s.isEmpty()){Toast.makeText(this,"أدخل قيمة",Toast.LENGTH_SHORT).show();return;}
            cb.accept(s);
        }).show();
    }
    void confirm(String msg,Runnable run) {
        new AlertDialog.Builder(this).setMessage(msg).setNegativeButton("إلغاء",null)
            .setPositiveButton("نعم", (d,w)->run.run()).show();
    }
    void home() {
        screen(config.optString("report_title","").trim().isEmpty()?"تقارير العمل اليومية":config.optString("report_title"));
        LinearLayout counts=new LinearLayout(this);counts.setOrientation(0);
        int people=0;HashSet<String> unique=new HashSet<>();
        for(String id:ids()){
            JSONArray staff=catFor(id).optJSONArray("people");
            if(staff!=null)for(int j=0;j<staff.length();j++)unique.add(staff.optString(j));
        }
        people=unique.size();
        statTile(counts,"التقارير",String.valueOf(reports.length()));
        statTile(counts,"الموظفون",String.valueOf(people));body.addView(counts);
        LinearLayout counts2=new LinearLayout(this);counts2.setOrientation(0);
        statTile(counts2,"الأقسام",String.valueOf(ids().size()));
        int fields=0;for(String id:ids()){
            JSONArray fs=catFor(id).optJSONArray("fields");if(fs!=null)fields+=fs.length();
        }
        statTile(counts2,"الحقول",String.valueOf(fields));body.addView(counts2);
        if(ids().isEmpty()) {
            label("أنشئ قسمك الأول، ثم أضف الحقول والأسماء التي تختارها.");
            button("＋ إنشاء القسم الأول",()->sections());
        }
        label("تقرير جديد");
        for(String id:ids())button("＋ "+title(id),()->reportForm(id));
        label("الإدارة السريعة");
        button("إدارة الموظفين",()->chooseCategory(false));
        button("إدارة الأقسام",()->sections());
        button("الحقول والقوائم المنسدلة",()->chooseCategory(true));
        label("Developer: Abu Omar | +966503276927");
        label("protech.2025.2026@gmail.com");
        label("النسخ الاحتياطي");
        button("مشاركة نسخة احتياطية JSON",()->shareData());
        button("استيراد إعدادات الكمبيوتر",()->importConfig());
    }
    void chooseCategory(boolean fields) {
        screen(fields?"إدارة الحقول والقوائم":"إدارة الموظفين");
        for(String id:ids())button(title(id),()->{if(fields) manageFields(id);else managePeople(id);});
    }
    void sections() {
        screen("إدارة الأقسام");
        button("＋ إضافة قسم جديد",()->editText("اسم القسم","",name->{
            String id="section_"+Long.toString(System.currentTimeMillis(),36);
            put(categories,id,cat(name,"",new JSONArray()));
            order.put(id);persist();sections();
        }));
        for(String id:ids()){
            label(title(id));
            button("تعديل مسمى الشخص أو الجهة",()->editText("أدخل المسمى الذي تختاره",
                catFor(id).optString("person_label",""),v->{put(catFor(id),"person_label",v);persist();sections();}));
            button("تعديل اسم القسم: "+title(id),()->editText("اسم القسم",title(id),v->{
                put(catFor(id),"label",v);persist();sections();
            }));
            button("حذف القسم: "+title(id),()->confirm("سيُحذف القسم من القوائم. لن تُحذف السجلات المحفوظة، ولكن القسم لن يظهر للإدخال.",()->{
                for(int i=0;i<reports.length();i++)if(id.equals(reports.optJSONObject(i).optString("category"))){
                    Toast.makeText(this,"القسم مرتبط بسجلات. لا يمكن حذفه.",Toast.LENGTH_LONG).show();return;
                }
                categories.remove(id);JSONArray next=new JSONArray();
                for(String k:ids())if(!k.equals(id))next.put(k);
                order=next;put(config,"category_order",next);persist();sections();
            }));
        }
    }
    void managePeople(String id) {
        JSONObject c=catFor(id);if(c==null){home();return;}
        screen("موظفو: "+title(id));
        JSONArray people=c.optJSONArray("people");
        if(people==null){people=new JSONArray();put(c,"people",people);}
        final JSONArray active=people;
        if(active.length()==0) label("القائمة فارغة. استخدم زر إضافة موظف.");
        button("＋ إضافة موظف",()->editText("اسم الموظف","",name->{
            for(int i=0;i<active.length();i++)if(active.optString(i).equalsIgnoreCase(name)){
                Toast.makeText(this,"الاسم موجود",Toast.LENGTH_SHORT).show();return;
            }
            active.put(name);persist();managePeople(id);
        }));
        for(int i=0;i<active.length();i++){
            final int index=i;String name=active.optString(i);
            button("تعديل: "+name,()->editText("تعديل اسم الموظف",name,v->{
                replace(active,index,v);persist();managePeople(id);
            }));
            button("إزالة من القائمة: "+name,()->confirm("إزالة "+name+" من الموظفين؟ ستبقى التقارير القديمة محفوظة.",()->{
                remove(active,index);persist();managePeople(id);
            }));
        }
    }
    void replace(JSONArray a,int index,Object v){try{a.put(index,v);}catch(Exception ignored){}}
    void remove(JSONArray a,int index){
        JSONArray next=new JSONArray();for(int j=0;j<a.length();j++)if(j!=index)next.put(a.opt(j));
        for(int j=a.length()-1;j>=0;j--)a.remove(j);
        for(int j=0;j<next.length();j++)a.put(next.opt(j));
    }
    void manageFields(String id) {
        JSONObject c=catFor(id);if(c==null){home();return;}
        screen("حقول: "+title(id));
        JSONArray fs=c.optJSONArray("fields");
        if(fs==null){fs=new JSONArray();put(c,"fields",fs);}
        JSONArray fields=fs;
        button("＋ إضافة حقل",()->fieldDialog(id,null,-1));
        for(int i=0;i<fields.length();i++){
            int index=i;JSONObject f=fields.optJSONObject(i);if(f==null)continue;
            String name=f.optString("label");
            button("تعديل: "+name+" • "+f.optString("kind","text"),()->fieldDialog(id,f,index));
            button("حذف الحقل: "+name,()->confirm("حذف الحقل من نموذج الإدخال؟ لن تُحذف قيمه في التقارير القديمة.",()->{
                remove(fields,index);persist();manageFields(id);
            }));
        }
    }
    void fieldDialog(String id,JSONObject existing,int at) {
        LinearLayout v=new LinearLayout(this);v.setOrientation(1);v.setPadding(22,10,22,5);
        EditText name=new EditText(this);name.setHint("اسم الحقل");name.setText(existing==null?"":existing.optString("label"));
        v.addView(name);
        Spinner type=new Spinner(this);String[] kinds={"رقم","نص","قائمة منسدلة"};
        ArrayAdapter<String> types=new ArrayAdapter<>(this,android.R.layout.simple_spinner_dropdown_item,kinds);
        type.setAdapter(types);
        if(existing!=null){String k=existing.optString("kind");type.setSelection(k.equals("number")?0:k.equals("select")?2:1);}
        v.addView(type);
        EditText opts=new EditText(this);opts.setHint("خيارات القائمة: خيار1، خيار2");opts.setSingleLine(false);
        if(existing!=null){JSONArray ar=existing.optJSONArray("options");if(ar!=null){StringBuilder sb=new StringBuilder();
            for(int i=0;i<ar.length();i++){if(i>0)sb.append("، ");sb.append(ar.optString(i));}opts.setText(sb.toString());}}
        v.addView(opts);
        new AlertDialog.Builder(this).setTitle(existing==null?"حقل جديد":"تعديل الحقل").setView(v)
          .setNegativeButton("إلغاء",null).setPositiveButton("حفظ",(d,w)->{
            String label=name.getText().toString().trim();if(label.isEmpty())return;
            JSONObject f=existing==null?new JSONObject():existing;
            if(existing==null)put(f,"id","custom_"+Long.toString(System.nanoTime(),36));
            put(f,"label",label);put(f,"kind",type.getSelectedItemPosition()==0?"number":type.getSelectedItemPosition()==2?"select":"text");
            JSONArray options=new JSONArray();String raw=opts.getText().toString();
            for(String o:raw.split("[،,\\n]"))if(!o.trim().isEmpty())options.put(o.trim());
            put(f,"options",options);
            if(existing==null)catFor(id).optJSONArray("fields").put(f);
            persist();manageFields(id);
        }).show();
    }
    void reportForm(String id) {
        JSONObject c=catFor(id);if(c==null){home();return;}
        screen("تقرير جديد • "+title(id));
        JSONArray names=c.optJSONArray("people");
        ArrayList<String> entries=new ArrayList<>();
        if(names!=null)for(int i=0;i<names.length();i++)entries.add(names.optString(i));
        if(entries.isEmpty()){label("أضف موظفًا أولًا من شاشة الإدارة");button("إدارة الموظفين",()->managePeople(id));return;}
        label(c.optString("person_label","الموظف"));
        Spinner people=new Spinner(this);
        people.setAdapter(new ArrayAdapter<String>(this,android.R.layout.simple_spinner_dropdown_item,entries));
        body.addView(people);
        JSONArray fields=c.optJSONArray("fields");if(fields==null)fields=new JSONArray();
        JSONObject values=new JSONObject();ArrayList<View> controls=new ArrayList<>();
        ArrayList<String> fieldIds=new ArrayList<>();
        for(int i=0;i<fields.length();i++) {
            JSONObject f=fields.optJSONObject(i);if(f==null)continue;
            String kind=f.optString("kind","text");String fid=f.optString("id");
            label(f.optString("label","الحقل"));
            if(kind.equals("select") || kind.equals("choice")){
                ArrayList<String> os=new ArrayList<>();os.add("");
                JSONArray options=f.optJSONArray("options");
                if(options!=null)for(int j=0;j<options.length();j++)os.add(options.optString(j));
                Spinner sp=new Spinner(this);sp.setAdapter(new ArrayAdapter<String>(this,android.R.layout.simple_spinner_dropdown_item,os));
                body.addView(sp);controls.add(sp);
            } else {
                EditText ed=new EditText(this);ed.setHint("القيمة");
                if(kind.equals("number"))ed.setInputType(InputType.TYPE_CLASS_NUMBER);
                body.addView(ed);controls.add(ed);
            }
            fieldIds.add(fid);
        }
        button("حفظ التقرير",()->{
            for(int i=0;i<fieldIds.size();i++){
                View view=controls.get(i);
                String v=view instanceof Spinner ? ((Spinner)view).getSelectedItem().toString() : ((EditText)view).getText().toString();
                put(values,fieldIds.get(i),v);
            }
            JSONObject row=new JSONObject();
            put(row,"id",UUID.randomUUID().toString());put(row,"date",stamp());put(row,"category",id);
            put(row,"category_name",title(id));put(row,"person",people.getSelectedItem().toString());put(row,"values",values);
            reports.put(row);persist();Toast.makeText(this,"تم حفظ التقرير",Toast.LENGTH_SHORT).show();history();
        });
    }
    void history() {
        screen("التقارير المحفوظة ("+reports.length()+")");
        EditText search=new EditText(this);search.setHint("بحث بالموظف أو القسم أو التاريخ");
        search.setSingleLine(true);search.setTextSize(15);
        body.addView(search);
        ArrayList<View> reportRows=new ArrayList<>();ArrayList<String> queries=new ArrayList<>();
        for(int i=reports.length()-1;i>=0;i--){
            JSONObject r=reports.optJSONObject(i);if(r==null)continue;
            String desc=r.optString("date")+" • "+r.optString("person")+" • "+r.optString("category_name");
            button(desc,()->{
                new AlertDialog.Builder(this).setTitle("تفاصيل التقرير").setMessage(r.toString()).setPositiveButton("إغلاق",null).show();
            });
            reportRows.add(body.getChildAt(body.getChildCount()-1));
            queries.add(desc.toLowerCase(java.util.Locale.ROOT));
        }
        search.addTextChangedListener(new TextWatcher(){
            public void beforeTextChanged(CharSequence q,int start,int count,int after){}
            public void onTextChanged(CharSequence q,int start,int before,int count){
                String term=q.toString().trim().toLowerCase(java.util.Locale.ROOT);
                for(int i=0;i<reportRows.size();i++)
                    reportRows.get(i).setVisibility(queries.get(i).contains(term)?View.VISIBLE:View.GONE);
            }
            public void afterTextChanged(Editable e){}
        });
        if(reports.length()==0)label("لا توجد تقارير محفوظة حتى الآن");
    }
    void shareData() {
        Intent in=new Intent(Intent.ACTION_SEND);in.setType("application/json");
        in.putExtra(Intent.EXTRA_TEXT,state.toString());
        startActivity(Intent.createChooser(in,"مشاركة نسخة احتياطية JSON"));
    }
    void importConfig() {
        Intent in=new Intent(Intent.ACTION_OPEN_DOCUMENT);
        in.addCategory(Intent.CATEGORY_OPENABLE);in.setType("application/json");
        startActivityForResult(in,importCode);
    }
    @Override protected void onActivityResult(int req,int result,Intent data) {
        super.onActivityResult(req,result,data);
        if(req!=importCode || result!=RESULT_OK || data==null || data.getData()==null)return;
        try(InputStream in=getContentResolver().openInputStream(data.getData())){
            ByteArrayOutputStream bout=new ByteArrayOutputStream();byte[] b=new byte[4096];int n;
            while((n=in.read(b))!=-1){bout.write(b,0,n);if(bout.size()>3000000)throw new Exception("الملف كبير");}
            JSONObject incoming=new JSONObject(bout.toString("UTF-8"));
            JSONObject imported=incoming.optJSONObject("config");
            if(imported==null)imported=incoming;
            if(imported.optJSONObject("categories")==null || imported.optJSONArray("category_order")==null)
                throw new Exception("ملف إعدادات غير متوافق");
            JSONObject fresh=imported;
            confirm("استبدال إعدادات الموظفين والأقسام والحقول؟ ستبقى التقارير المحفوظة.",()->{
                put(state,"config",fresh);config=fresh;categories=config.optJSONObject("categories");
                order=config.optJSONArray("category_order");persist();home();
            });
        } catch(Exception ex) {new AlertDialog.Builder(this).setMessage("تعذر الاستيراد: "+ex.getMessage()).setPositiveButton("حسنًا",null).show();}
    }
    @Override public void onBackPressed(){home();}
}