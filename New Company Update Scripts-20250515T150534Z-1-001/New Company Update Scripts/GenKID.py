class UniqueIDGenerator:
    def __init__(self, prefix="INA", start=1, end=99999):
        """
        Initialize the ID generator.
        :param prefix: The constant prefix, e.g., "INA".
        :param start: The starting number for IDs.
        :param end: The maximum number for IDs.
        """
        self.prefix = prefix
        self.start = start
        self.end = end

    def generate_ids(self, series_start, series_end):
        """
        Generate a series of unique IDs.
        :param series_start: Start of the desired series (inclusive).
        :param series_end: End of the desired series (inclusive).
        :return: A list of generated IDs.
        """
        if series_start < self.start or series_end > self.end:
            raise ValueError(f"Series range must be between {self.start} and {self.end}.")
        
        if series_start > series_end:
            raise ValueError("Start of series must not be greater than the end.")
        
        ids = [f"{self.prefix}{str(i).zfill(5)}" for i in range(series_start, series_end + 1)]
        return ids


# Example usage with user input
if __name__ == "__main__":
    # Initialize the generator
    generator = UniqueIDGenerator()
    
    print("Welcome to the Unique ID Generator!")
    print(f"The prefix is '{generator.prefix}', and IDs range from {generator.start} to {generator.end}.")
    
    try:
        # Get user input for the start and end of the range
        series_start = int(input("Enter the start of the series (e.g., 1): "))
        series_end = int(input("Enter the end of the series (e.g., 5): "))
        
        # Generate the IDs
        ids = generator.generate_ids(series_start, series_end)
        print("\nGenerated IDs:")
        for _id in ids:
            print(_id)
    
    except ValueError as e:
        print(f"Error: {e}")
