package com.spring.excel.repository;

import java.io.Serializable;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.repository.NoRepositoryBean;

import com.spring.excel.model.ExcelSheet;

@NoRepositoryBean // important while using the generic Repo and it should not create an entity for ExcelSheet abstract class.
public interface CommonRepo<T extends ExcelSheet, ID extends Serializable> extends JpaRepository<T,ID> {

}
