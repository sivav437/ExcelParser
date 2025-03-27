package com.spring.excel.model;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name="product")
public class Product extends ExcelSheet {
	
	@Id
	private int product_id;
	
	private String product_name;
	
	private int product_price;

	public int getProduct_id() {
		return product_id;
	}

	
//	public Product clone() throws CloneNotSupportedException {
//		Product cloned=(Product) super.clone();
//		return cloned;
//	}
	
	public void setProduct_id(int product_id) {
		this.product_id = product_id;
	}

	public String getProduct_name() {
		return product_name;
	}

	public void setProduct_name(String product_name) {
		this.product_name = product_name;
	}

	public int getProduct_price() {
		return product_price;
	}

	public void setProduct_price(int product_price) {
		this.product_price = product_price;
	}

	public Product(int product_id, String product_name, int product_price) {
		super();
		this.product_id = product_id;
		this.product_name = product_name;
		this.product_price = product_price;
	}

	public Product() {
		super();
	}

	@Override
	public String toString() {
		return "Product [product_id=" + product_id + ", product_name=" + product_name + ", product_price="
				+ product_price + "]";
	}
	
	

}
