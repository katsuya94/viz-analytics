class AddValueToData < ActiveRecord::Migration
  def change
    add_column :data, :value, :float
  end
end
