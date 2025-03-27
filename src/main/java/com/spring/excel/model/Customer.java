package com.spring.excel.model;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name="customer")
public class Customer extends ExcelSheet {
	
	
	@Id
//	@GeneratedValue(strategy=GenerationType.IDENTITY)
	private int customer_id;
	
	private String customer_name;
	
	private String customer_password;

	public int getCustomer_id() {
		return customer_id;
	}
	
	// not required coz throwing errors
	
//	public Customer clone() throws CloneNotSupportedException {
//		Customer cloned=(Customer) super.clone();
//		return cloned;
//	}

	public void setCustomer_id(int customer_id) {
		this.customer_id = customer_id;
	}

	public String getCustomer_name() {
		return customer_name;
	}

	public void setCustomer_name(String customer_name) {
		this.customer_name = customer_name;
	}

	public String getCustomer_password() {
		return customer_password;
	}

	public void setCustomer_password(String customer_password) {
		this.customer_password = customer_password;
	}
	
	

	public Customer() {
		super();
	}

	public Customer(int customer_id, String customer_name, String customer_password) {
		super();
		this.customer_id = customer_id;
		this.customer_name = customer_name;
		this.customer_password = customer_password;
	}

	@Override
	public String toString() {
		return "Customer [customer_id=" + customer_id + ", customer_name=" + customer_name + ", customer_password="
				+ customer_password + "]";
	}
	
	
	

}
