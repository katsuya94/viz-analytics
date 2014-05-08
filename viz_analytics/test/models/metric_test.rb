require 'test_helper'

class MetricTest < ActiveSupport::TestCase
	test "Raw Metrics" do
		puts metrics(:bounce_rate).get_datum(companies(:buzzfeed))
	end
end
