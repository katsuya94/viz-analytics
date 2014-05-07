class CreateCompanies < ActiveRecord::Migration
  def change
    create_table :companies do |t|
    	t.string :name
      t.timestamps
    end
    add_reference :data, :company
  end
end
