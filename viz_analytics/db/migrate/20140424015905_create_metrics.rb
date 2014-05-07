class CreateMetrics < ActiveRecord::Migration
  def change
    create_table :metrics do |t|

      t.timestamps
      t.float :value

    end

    add_column :metric, :lifetime, :integer
    
  end
end
