ActiveRecord::Schema.define(:version => 0) do

  create_table :users, :force => true do |t|
    t.column :name, :string
    t.column :age, :integer
  end
end
