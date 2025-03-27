package com.spring.excel.model;

import jakarta.persistence.MappedSuperclass;

@MappedSuperclass
//@Entity  // used for managing generic JPA Repo 
//@Inheritance(strategy = InheritanceType.TABLE_PER_CLASS)
public abstract class  ExcelSheet { 
	
	// should obly use abstract class for parent entity hence JpaRepo only works with concrete classes.

}
