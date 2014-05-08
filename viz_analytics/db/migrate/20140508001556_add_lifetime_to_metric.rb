class AddLifetimeToMetric < ActiveRecord::Migration
  def change
    add_column :metrics, :lifetime, :integer
  end
end
