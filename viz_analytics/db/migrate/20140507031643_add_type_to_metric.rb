class AddTypeToMetric < ActiveRecord::Migration
	def change
		add_column :metrics, :type, :string
	end
end
