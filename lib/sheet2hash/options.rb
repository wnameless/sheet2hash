module Sheet2hash
  # Sheet2hash::Options processes the options of Sheet2hash.
  # @author Wei-Ming Wu
  module Options
    # Regulates options of Sheet2hash.
    def process_options opts
      regulate_options opts
      opts
    end
    
    private
    
    def regulate_options opts
      opts[:keep_row] = Array(opts[:keep_row]) if opts[:keep_row]
      opts[:skip_row] = Array(opts[:skip_row]) if opts[:skip_row]
      opts[:keep_col] = Array(opts[:keep_col]) if opts[:keep_col]
      opts[:skip_col] = Array(opts[:skip_col]) if opts[:skip_col]
    end
  end
end
