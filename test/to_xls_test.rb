require 'test_helper'

class User < ActiveRecord::Base
  def is_old?
    age > 22
  end
end

class ToXlsTest < TestCaseClass
  def xls_doc(content)
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Workbook xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\" xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">#{content}</Workbook>"
  end
  
  def worksheet(name, content)
    "<Worksheet ss:Name=\"#{name}\">#{content}</Worksheet>"
  end
  
  def worksheet_doc(worksheet_name, content)
    xls_doc(worksheet(worksheet_name, content))
  end
  
  def cell(type, content)
    "<Cell><Data ss:Type=\"#{type}\">#{content}</Data></Cell>"
  end
  
  def row(*cells)
    "<Row>#{cells.join}</Row>"
  end
  
  def table(*rows)
    "<Table>#{rows.join}</Table>"
  end
  
  # *****
  def setup
    @users = []
    @users << User.new(:id => 1, :name => 'Ary', :age => 24)
    @users << User.new(:id => 2, :name => 'Nati', :age => 21)
  end

  def test_with_empty_array
    assert_equal( worksheet_doc("Sheet1", table), [].to_xls )
  end
  
  def test_with_no_options
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Age"), cell("String", "Name")),
      row(cell("Number", 24), cell("String", "Ary")),
      row(cell("Number", 21), cell("String", "Nati")))),
      @users.to_xls
    )
  end
  
  def test_with_no_headers
    assert_equal(worksheet_doc("Sheet1", table(
      row(cell("Number", 24), cell("String", "Ary")),
      row(cell("Number", 21), cell("String", "Nati")))),
      @users.to_xls(:headers => false)
    ) 
  end
  
  def test_with_only
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Name")),
      row(cell("String", "Ary")),
      row(cell("String", "Nati")))),
      @users.to_xls(:only => :name)
    )
  end
  
  def test_with_empty_only
    assert_equal( worksheet_doc("Sheet1", table), @users.to_xls(:only => "") )
  end
  
  def test_with_only_and_wrong_column_names
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Name")),
      row(cell("String", "Ary")),
      row(cell("String", "Nati")))),
      @users.to_xls(:only => [:name, :yoyo])
    )
  end
  
  def test_with_except
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Age")),
      row(cell("Number", 24)),
      row(cell("Number", 21)))),
      @users.to_xls(:except => [:id, :name])
    )
  end
  
  def test_with_except_and_only_should_listen_to_only
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Name")),
      row(cell("String", "Ary")),
      row(cell("String", "Nati")))),
      @users.to_xls(:except => [:id, :name], :only => :name)
    )
  end
  
  def test_with_methods
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Age"), cell("String", "Name"), cell("String", "Is old?")),
      row(cell("Number", 24), cell("String", "Ary"), cell("String", "true")),
      row(cell("Number", 21), cell("String", "Nati"), cell("String", "false")))),
      @users.to_xls(:methods => [:is_old?])
    )
  end
  
  def test_with_i18n
    old_locale = I18n.locale
    I18n.locale = "pt-PT"
    I18n.backend.store_translations("pt-PT", {
      "activerecord" => {
        "attributes" => {
          "user" => {
            "name" => "Nome",
            "age" => "Idade"
          }
        }
      }
    })
    assert_equal( worksheet_doc("Sheet1", table(
      row(cell("String", "Idade"), cell("String", "Nome")),
      row(cell("Number", 24), cell("String", "Ary")),
      row(cell("Number", 21), cell("String", "Nati")))),
      @users.to_xls
    )
  ensure
    I18n.locale = old_locale
  end
  
end
