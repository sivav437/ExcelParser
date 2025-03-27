package com.spring.excel.repository;

import org.springframework.stereotype.Repository;

import com.spring.excel.model.Product;

@Repository("productRepo")   // use name  which is suggested
public interface ProductRepo extends CommonRepo<Product, Integer> {

}
