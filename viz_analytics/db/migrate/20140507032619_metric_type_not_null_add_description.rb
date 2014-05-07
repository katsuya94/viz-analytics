class MetricTypeNotNullAddDescription < ActiveRecord::Migration
  def change
  	change_column :metrics, :type, :string, :null => false
  	add_column :metrics, :description, :string
  end
end
