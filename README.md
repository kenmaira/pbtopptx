
# PB to PPTX

A script to generate PowerPoint slides from Productboard release features.

## Features

- Fetches feature details from Productboard API.
- Inserts images and descriptions into slides using a PowerPoint template.
- Maintains aspect ratio for images.
- Supports filtering features by owner email.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/pbtopptx.git
   cd pbtopptx
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Add your API token in the `pbtopptx.py` file:
   ```python
   headers = {
       "accept": "application/json",
       "X-Version": "1",
       "authorization": "Bearer YOUR_API_TOKEN"
   }
   ```

## Usage

Run the script with:
```bash
python pbtopptx.py <release_id> [--owner_email <email>]
```

Example:
```bash
python pbtopptx.py 92bc2174-0096-4ac5-bb69-d1ee9cb2825f --owner_email john.doe@example.com
```

## Requirements

- Python 3.7 or higher
- Libraries listed in `requirements.txt`

## Contributing

1. Fork the repository.
2. Create a new branch: `git checkout -b feature-name`.
3. Make your changes and commit them: `git commit -m 'Add new feature'`.
4. Push to the branch: `git push origin feature-name`.
5. Create a pull request.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
