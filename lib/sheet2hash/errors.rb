module Sheet2hash
  # Sheet2hash::Errors defines all errors of Sheet2hash.
  # @author Wei-Ming Wu
  module Errors
    # SheetNotFoundError is an Exception.
    class SheetNotFoundError < Exception ; end
    # InvalidHeaderError is an Exception.
    class InvalidHeaderError < Exception ; end
  end
end
