package com.dbPoiXlsx;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import org.junit.Test;

public class jdbcConfig {
	
	@Test
	public Connection connectPostgre() {
		try {
			Class.forName("org.postgresql.Driver");
		} catch (Exception e) {
			System.out.println("Wrong PostgreSQL JDBC Driver? ");
		    e.printStackTrace();
		    return null;
		}
		
		System.out.println("postgresql driver registered");
		
		Connection connection = null;
		
		try {
			connection = DriverManager.getConnection("jdbc:postgresql://125.77.26.134:5432/fjbak", "postgres", "********");
		} catch (SQLException e) {
			System.out.println("connect db failed");
			e.printStackTrace();
			return null;
		}
		
		return connection;
		
	}
}
