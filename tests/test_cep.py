import unittest
from unittest.mock import patch, MagicMock
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

class TestCEPFunctions(unittest.TestCase):
    
    @patch('requests.get')
    def test_get_cep_from_address_success(self, mock_get):
        """Test successful CEP lookup from address"""
        from src.cep import get_cep_from_address
        
        # Mock successful response
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'cep': '01305-000',
            'logradouro': 'Rua Augusta',
            'bairro': 'Consolação',
            'localidade': 'São Paulo',
            'uf': 'SP'
        }
        mock_get.return_value = mock_response
        
        result = get_cep_from_address('Rua Augusta, Sao Paulo')
        self.assertEqual(result, '01305-000')
        mock_get.assert_called_once()
    
    @patch('requests.get')
    def test_get_cep_from_address_not_found(self, mock_get):
        """Test CEP lookup when address not found"""
        from src.cep import get_cep_from_address
        
        # Mock response with no results
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = []
        mock_get.return_value = mock_response
        
        result = get_cep_from_address('Endereço Inexistente')
        self.assertIsNone(result)
    
    @patch('requests.get')
    def test_get_cep_from_address_error(self, mock_get):
        """Test CEP lookup when API returns error"""
        from src.cep import get_cep_from_address
        
        # Mock error response
        mock_response = MagicMock()
        mock_response.status_code = 400
        mock_get.return_value = mock_response
        
        result = get_cep_from_address('Qualquer endereço')
        self.assertIsNone(result)
    
    @patch('requests.get')
    def test_get_cep_from_address_timeout(self, mock_get):
        """Test CEP lookup timeout handling"""
        from src.cep import get_cep_from_address
        import requests
        
        # Mock timeout exception
        mock_get.side_effect = requests.exceptions.Timeout()
        
        result = get_cep_from_address('Rua Augusta, Sao Paulo')
        self.assertIsNone(result)
    
    @patch('requests.get')
    def test_get_cep_from_address_connection_error(self, mock_get):
        """Test CEP lookup connection error handling"""
        from src.cep import get_cep_from_address
        import requests
        
        # Mock connection error
        mock_get.side_effect = requests.exceptions.ConnectionError()
        
        result = get_cep_from_address('Rua Augusta, Sao Paulo')
        self.assertIsNone(result)
    
    def test_format_address(self):
        """Test address formatting function"""
        from src.cep import format_address
        
        # Test basic formatting
        formatted = format_address('rua augusta, sao paulo')
        self.assertEqual(formatted, 'Rua Augusta, Sao Paulo')
        
        # Test with extra spaces
        formatted = format_address('  rua augusta,  sao paulo  ')
        self.assertEqual(formatted, 'Rua Augusta, Sao Paulo')
    
    @patch('requests.get')
    def test_get_address_from_cep_success(self, mock_get):
        """Test successful address lookup from CEP"""
        from src.cep import get_address_from_cep
        
        # Mock successful response
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'cep': '01305-000',
            'logradouro': 'Rua Augusta',
            'complemento': '',
            'bairro': 'Consolação',
            'localidade': 'São Paulo',
            'uf': 'SP',
            'ibge': '3550308',
            'gia': '1004',
            'ddd': '11',
            'siafi': '7107'
        }
        mock_get.return_value = mock_response
        
        result = get_address_from_cep('01305-000')
        self.assertIsNotNone(result)
        self.assertEqual(result['cep'], '01305-000')
        self.assertEqual(result['logradouro'], 'Rua Augusta')
        self.assertEqual(result['localidade'], 'São Paulo')
    
    @patch('requests.get')
    def test_get_address_from_cep_invalid_format(self, mock_get):
        """Test address lookup with invalid CEP format"""
        from src.cep import get_address_from_cep
        
        # Test various invalid formats
        result = get_address_from_cep('12345')  # Too short
        self.assertIsNone(result)
        
        result = get_address_from_cep('123456789')  # Too long
        self.assertIsNone(result)
        
        result = get_address_from_cep('abcde-fgh')  # Non-numeric
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()
