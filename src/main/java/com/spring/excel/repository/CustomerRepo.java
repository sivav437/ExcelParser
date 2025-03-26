package com.spring.excel.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.spring.excel.model.Customer;

@Repository
public interface CustomerRepo extends JpaRepository<Customer, Integer>,CommonRepo {

}
