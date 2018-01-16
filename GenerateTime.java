package com.dbPoiXlsx;

import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.junit.Test;

public class GenerateTime {
    public Timestamp generateTs(String str) throws ParseException {
        SimpleDateFormat sdf;//小写的mm表示的是分钟
        Date date;
        try {
            sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            date = sdf.parse(str);
        } catch (ParseException e) {
            sdf=new SimpleDateFormat("yyyy-MM-dd");
            date = sdf.parse(str);
        }

        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String time = df.format(date);
        Timestamp ts = Timestamp.valueOf(time);
        return ts;
    }

    @Test
    public void test() throws ParseException {
        this.generateTs("2018-01-09");
    }
}
