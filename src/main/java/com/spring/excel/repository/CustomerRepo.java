package com.spring.excel.repository;

import org.springframework.stereotype.Repository;

import com.spring.excel.model.Customer;

@Repository("customerRepo") // use name  which is suggested
public interface CustomerRepo extends CommonRepo<Customer, Integer> {

}
