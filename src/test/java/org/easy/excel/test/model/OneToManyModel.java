package org.easy.excel.test.model;

import java.util.List;

/**
 * 一对多例子Model
 * @author lisuo
 *
 */
public class OneToManyModel {

	private String studentName;
	private List<BookModel> books;

	public String getStudentName() {
		return studentName;
	}

	public void setStudentName(String studentName) {
		this.studentName = studentName;
	}

	public List<BookModel> getBooks() {
		return books;
	}

	public void setBooks(List<BookModel> books) {
		this.books = books;
	}

	@Override
	public String toString() {
		return "OneToManyModel [studentName=" + studentName + ", books=" + books + "]";
	}

	
}
