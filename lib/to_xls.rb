class Array
  
  def to_xls(options = {})
    output = '<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office"><Worksheet ss:Name="Sheet1"><Table>'
    
    if self.any?
      instance = self.first
      attributes = instance.attributes.keys.sort.map { |c| c.to_sym }
      
      if options[:only]
        # the "& attributes" get rid of invalid columns
        columns = Array(options[:only]) & attributes
      else
        columns = attributes - Array(options[:except])
      end

      columns += options[:methods].to_a
    
      if columns.any?
        unless options[:headers] == false
          output << "<Row>"
          columns.each do |column|
            output << "<Cell><Data ss:Type=\"String\">#{instance.class.human_attribute_name(column)}</Data></Cell>"
          end
          output << "</Row>"
        end    

        self.each do |item|
          output << "<Row>"
          columns.each do |column|
            value = item.send(column)
            output << "<Cell><Data ss:Type=\"#{value.kind_of?(Integer) ? 'Number' : 'String'}\">#{value}</Data></Cell>"
          end
          output << "</Row>"
        end
      end
      
    end
    
    output << '</Table></Worksheet></Workbook>'
  end
  
end
