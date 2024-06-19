# WooCommerce Product Management Script

This Python script enables you to efficiently manage products in your WooCommerce store by uploading them from an Excel file, updating existing products, and handling categories and images seamlessly via the WooCommerce API.

## Prerequisites

Before running the script, ensure you have the following:

1. **Python Environment**: Python 3.7 or higher installed on your machine.
   
2. **Dependencies**: Install required Python packages using `pip`:
   ```bash
   pip install requests openpyxl woocommerce ratelimit
   ```

3. **WooCommerce API Credentials**: You need to obtain your WooCommerce store's API credentials:
   - **Consumer Key**: This acts as the username for authentication with the WooCommerce API.
   - **Consumer Secret**: This is the password/token for authentication with the WooCommerce API.

   To get these credentials:
   - Log in to your WooCommerce store's WordPress admin area.
   - Navigate to **WooCommerce > Settings > Advanced > REST API**.
   - Click on **Add Key** under **Keys/Apps**.
   - Provide a **Description** (e.g., "Script API Access").
   - Set the **User** to a valid user (typically admin).
   - Choose **Permissions** (`Read/Write` recommended for full functionality).
   - Click **Generate API Key**.
   - Copy the generated **Consumer Key** and **Consumer Secret**. **Note:** These are shown only once. If you lose them, you'll need to regenerate the key.

4. **Environment Variables (Optional)**: If your images are stored locally and not directly accessible via URLs, set the path to your local image directory using the `IMAGES_PATH` environment variable.

## Configuration

Update the `config` dictionary in the script (`main.py`) with your WooCommerce API credentials and store URL:

```python
config = {
    'WOOCOMMERCE_URL': 'your_woocommerce_url',
    'WOOCOMMERCE_CONSUMER_KEY': 'your_consumer_key',
    'WOOCOMMERCE_CONSUMER_SECRET': 'your_consumer_secret'
}
```

## Usage

1. **Generate a Products File Template**:
   - Included in the repository is a script that generates a products file template (`generate_products_template.py`). You can run this script to create a sample Excel file with random data:
     ```bash
     python generate_products_template.py
     ```
   - This will generate a file named `products.xlsx` with the required columns and some sample data, which you can use as a template.

2. **Prepare Excel File**:
   - Create or edit the Excel file (`products.xlsx`) with the following columns (in order):
     - Category
     - Subcategory
     - Image Path (if uploading images)
     - Product Name
     - Description
     - Price
     - Reference (SKU)
     - Brand
     - Site Status (`SIM` for publish, `NÃO` for draft)

3. **Run the Script**:
   - Execute the script `main.py`:
     ```bash
     python main.py
     ```
   - The script will:
     - Read the `products.xlsx` file.
     - Upload products to WooCommerce, update the Excel file with image URLs, and generate an `updated_products.json` file with product details.

4. **Logging and Output**:
   - Logs (`script.log`) will be generated in the same directory, providing information and error messages encountered during the script's execution.
   - Check `updated_products.json` for details of the products that were updated.

## Additional Notes

- **Error Handling**: If any errors occur during product upload or update, detailed error messages will be logged to `script.log`.
- **Rate Limiting**: The script incorporates rate limiting to avoid exceeding WooCommerce API rate limits, ensuring stable operation.

