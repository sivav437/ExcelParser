package com.spring.excel.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.spring.excel.model.Product;

@Repository
public interface ProductRepo extends JpaRepository<Product, Integer>,CommonRepo {

}
