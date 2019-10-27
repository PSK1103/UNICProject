package com.unic.unicproject;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import com.google.firebase.firestore.FirebaseFirestore;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    private Button btnUpload,btnRetrieve,btnSelect;
    private FirebaseFirestore db;
    private HSSFWorkbook workbook;
    private File excelFile;
    private int noOfSheets;
    private static final int EXCEL_FILE = 101;
    public MainActivity() {
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        btnSelect = findViewById(R.id.btnSelect);
        btnSelect.setOnClickListener(this);
        btnUpload = findViewById(R.id.btnUpload);
        btnUpload.setOnClickListener(this);
        btnRetrieve = findViewById(R.id.btnRetrieve);
        btnRetrieve.setOnClickListener(this);

        db = FirebaseFirestore.getInstance();
    }
    public void onClick(View view){
        switch (view.getId()){
            case R.id.btnSelect:
                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.setType("*/*");
                startActivityForResult(intent,EXCEL_FILE);
                break;
            case R.id.btnUpload:
                getData();
                break;
        }
    }


    public void getData(){
        try {
            FileInputStream fis = new FileInputStream(excelFile);
            //Workbook workbook1 = WorkbookFactory.create(fis);
            //XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
            workbook= new HSSFWorkbook(fis);
            noOfSheets = workbook.getNumberOfSheets();
            HSSFSheet[] sheet = new HSSFSheet[noOfSheets];
            for(int i=0;i<noOfSheets;i++){
                sheet[i]=workbook.getSheetAt(i);
                ArrayList<String> Header = new ArrayList<>();
                Iterator<Row> rowIterator = sheet[i].rowIterator();
                if(rowIterator.hasNext()){
                    HSSFRow header = (HSSFRow)rowIterator.next();
                    Iterator<Cell> headIterator = header.cellIterator();
                    do {
                        HSSFCell cell = (HSSFCell)headIterator.next();
                        Header.add(cell.toString());
                        Toast.makeText(this, cell.toString(), Toast.LENGTH_SHORT).show();
                    }while(headIterator.hasNext());
                    do {
                        HashMap<String,String> dataMap = new HashMap<>();
                        HSSFRow row = (HSSFRow) rowIterator.next();
                        Iterator<Cell> cellIterator = row.cellIterator();
                        int j=0;
                        do {
                            HSSFCell cell = (HSSFCell) cellIterator.next();
                            dataMap.put(Header.get(j),cell.toString());
                            j++;
                        }while(cellIterator.hasNext());

                        db.collection("products").add(dataMap);
                    }while(rowIterator.hasNext());

                }
            }
        } catch (Exception e){
            Toast.makeText(this, e.toString(), Toast.LENGTH_SHORT).show();
            e.printStackTrace();
        }

    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if(requestCode==EXCEL_FILE&&resultCode==RESULT_OK){
            Uri uri = data.getData();
            getFilePath(uri);
        }
    }
    private void getFilePath(@NotNull Uri uri){
        if(isDocumentUri(uri)) {
            String loc = uri.getPath().split("/")[2].split(":")[0];
            String[] trim = uri.getPath().split(":");
            if(loc.equals("home"))
                excelFile = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), trim[1]);
            else if (loc.equals("primary"))
                excelFile = new File("/storage/emulated/0/"+trim[1]);
        } else if(isDownloadsUri(uri)){
            Toast.makeText(this, uri.getPath(), Toast.LENGTH_SHORT).show();
            String[] trim = uri.getPath().split(":");
            excelFile = new File(trim[1]);
        } else {
            //Toast.makeText(this, uri.getPath(), Toast.LENGTH_SHORT).show();
        }
    }

    private boolean isDocumentUri(@NotNull Uri uri){
        return "com.android.externalstorage.documents".equals(uri.getAuthority());
    }
    private boolean isDownloadsUri(@NotNull Uri uri){
        return "com.android.providers.downloads.documents".equals(uri.getAuthority());
    }
}
