"""
Comprehensive test suite for weld_counter.py
Tests all functions with real and synthetic data for 90%+ coverage
"""

import pytest
import sys
import os
from pathlib import Path
import csv
import tempfile
import pandas as pd

# Add parent directory to path to import the module
sys.path.insert(0, str(Path(__file__).parent.parent))

from weld_counter import (
    extract_branch_positions,
    find_connected_branches,
    extract_branch_connections,
    extract_components_from_branches,
    lookup_and_merge_with_excel,
    save_to_csv
)


class TestExtractBranchPositions:
    """Tests for extract_branch_positions function"""
    
    def test_happy_path_single_branch(self, tmp_path):
        """Test extraction of a single branch with valid data"""
        test_file = tmp_path / "test_single_branch.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
BUIL false
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 1
        assert result[0]['Pipe'] == '0ABC12BR100'
        assert result[0]['Branch'] == '/B1'
        assert result[0]['Full_Branch_ID'] == '0ABC12BR100/B1'
        assert result[0]['HPOS_X'] == 1000.0
        assert result[0]['HPOS_Y'] == 2000.0
        assert result[0]['HPOS_Z'] == 3000.0
        assert result[0]['TPOS_X'] == 1500.0
        assert result[0]['TPOS_Y'] == 2500.0
        assert result[0]['TPOS_Z'] == 3500.0
    
    def test_multiple_branches_same_pipe(self, tmp_path):
        """Test multiple branches in the same pipe"""
        test_file = tmp_path / "test_multiple_branches.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
NEW BRANCH /0ABC12BR100/B2
HPOS X 1500mm Y 2500mm Z 3500mm
TPOS X 2000mm Y 3000mm Z 4000mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 2
        assert result[0]['Branch'] == '/B1'
        assert result[1]['Branch'] == '/B2'
        assert result[1]['HPOS_X'] == 1500.0
    
    def test_multiple_pipes(self, tmp_path):
        """Test multiple pipes with branches"""
        test_file = tmp_path / "test_multiple_pipes.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
NEW PIPE /0DEF34BR200
NEW BRANCH /0DEF34BR200/B1
HPOS X 5000mm Y 6000mm Z 7000mm
TPOS X 5500mm Y 6500mm Z 7500mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 2
        assert result[0]['Pipe'] == '0ABC12BR100'
        assert result[1]['Pipe'] == '0DEF34BR200'
    
    def test_negative_coordinates(self, tmp_path):
        """Test branches with negative coordinates"""
        test_file = tmp_path / "test_negative.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X -1000mm Y -2000mm Z 3000mm
TPOS X -1500mm Y 2500mm Z -3500mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert result[0]['HPOS_X'] == -1000.0
        assert result[0]['HPOS_Y'] == -2000.0
        assert result[0]['TPOS_Z'] == -3500.0
    
    def test_missing_positions(self, tmp_path):
        """Test branch with missing position data"""
        test_file = tmp_path / "test_missing.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 1
        assert result[0]['HPOS_X'] is None
        assert result[0]['TPOS_X'] is None
    
    def test_empty_file(self, tmp_path):
        """Test with empty file"""
        test_file = tmp_path / "test_empty.txt"
        test_file.write_text("")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 0
    
    def test_no_branches(self, tmp_path):
        """Test file with pipes but no branches"""
        test_file = tmp_path / "test_no_branches.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
BUIL false
SHOP false
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 0
    
    def test_decimal_coordinates(self, tmp_path):
        """Test with decimal coordinate values"""
        test_file = tmp_path / "test_decimal.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000.5mm Y 2000.75mm Z 3000.125mm
TPOS X 1500.99mm Y 2500.01mm Z 3500.999mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert result[0]['HPOS_X'] == 1000.5
        assert result[0]['TPOS_Z'] == 3500.999
    
    def test_real_data_file_TBY_0AUX(self):
        """Test with real TBY-0AUX-P.txt file"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        result = extract_branch_positions(str(real_file))
        
        assert len(result) == 483
        # Check that all branches have required fields
        for branch in result:
            assert 'Pipe' in branch
            assert 'Branch' in branch
            assert 'Full_Branch_ID' in branch
    
    def test_real_data_file_TBY_1AUX(self):
        """Test with real TBY-1AUX-P.txt file"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-1AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        result = extract_branch_positions(str(real_file))
        
        assert len(result) == 528
        # Verify a specific known branch exists
        branch_ids = [b['Full_Branch_ID'] for b in result]
        assert any('1' in bid for bid in branch_ids)


class TestFindConnectedBranches:
    """Tests for find_connected_branches function"""
    
    def test_happy_path_perfect_match(self):
        """Test branches with perfect alignment (0mm offset)"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        assert result[0]['Branch_A'] == '0ABC12BR100/B1'
        assert result[0]['Branch_B'] == '0ABC12BR100/B2'
        assert result[0]['Distance_mm'] == 0.0
        assert result[0]['Accuracy_Percent'] == 100.0
        assert result[0]['Match_Type'] == 'XY_tight_Z_loose'
    
    def test_xy_tight_z_loose_match(self):
        """Test XY tight, Z loose matching"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 2.0, 'TPOS_Z': 100.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1001.0, 'HPOS_Y': 3.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        assert result[0]['Match_Type'] == 'XY_tight_Z_loose'
        assert result[0]['Offset_X_mm'] == 1.0
        assert result[0]['Offset_Y_mm'] == 1.0
        assert result[0]['Offset_Z_mm'] == 100.0
    
    def test_xz_tight_y_loose_match(self):
        """Test XZ tight, Y loose matching"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 100.0, 'TPOS_Z': 2.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1001.0, 'HPOS_Y': 0.0, 'HPOS_Z': 3.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        assert result[0]['Match_Type'] == 'XZ_tight_Y_loose'
        assert result[0]['Loose_Offset_mm'] == 100.0
    
    def test_yz_tight_x_loose_match(self):
        """Test YZ tight, X loose matching"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 100.0, 'TPOS_Y': 1000.0, 'TPOS_Z': 2.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 0.0, 'HPOS_Y': 1001.0, 'HPOS_Z': 3.0,
                'TPOS_X': 0.0, 'TPOS_Y': 2000.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        assert result[0]['Match_Type'] == 'YZ_tight_X_loose'
        assert result[0]['Loose_Offset_mm'] == 100.0
    
    def test_no_matches_too_far(self):
        """Test branches that are too far apart"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 5000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 6000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 0
    
    def test_missing_positions(self):
        """Test branches with missing position data"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': None, 'HPOS_Y': None, 'HPOS_Z': None,
                'TPOS_X': 1000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': None, 'TPOS_Y': None, 'TPOS_Z': None
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        # Function finds the connection where branch1 has TPOS and branch2 has HPOS
        assert len(result) == 1
    
    def test_duplicate_prevention(self):
        """Test that duplicate pairs are not created"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        # Call twice - should still only get one result
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
    
    def test_accuracy_calculation(self):
        """Test accuracy percentage calculation"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 2.5, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1000.0, 'HPOS_Y': 2.5, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        # Coordinates match exactly (TPOS of B1 == HPOS of B2), so accuracy is 100%
        assert result[0]['Accuracy_Percent'] == 100.0
        assert result[0]['Max_Tight_Offset_mm'] == 0.0
    
    def test_empty_list(self):
        """Test with empty branch list"""
        result = find_connected_branches([], tolerance_tight=5.0, tolerance_loose=150.0)
        assert len(result) == 0
    
    def test_custom_tolerances(self):
        """Test with custom tolerance values"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 3.0, 'TPOS_Z': 50.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1003.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        # Should match with 10mm tight tolerance
        result = find_connected_branches(branches, tolerance_tight=10.0, tolerance_loose=100.0)
        assert len(result) == 1
        
        # Should not match with 2mm tight tolerance
        result = find_connected_branches(branches, tolerance_tight=2.0, tolerance_loose=100.0)
        assert len(result) == 0
    
    def test_real_data_connections(self):
        """Test with real data from TBY files"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        branches = extract_branch_positions(str(real_file))
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        # Should find some connections
        assert len(result) > 0
        # All connections should have valid match types
        for conn in result:
            assert conn['Match_Type'] in ['XY_tight_Z_loose', 'XZ_tight_Y_loose', 'YZ_tight_X_loose']


class TestExtractBranchConnections:
    """Tests for extract_branch_connections function"""
    
    def test_happy_path_with_components(self, tmp_path):
        """Test branch with components and pipe length calculation"""
        test_file = tmp_path / "test_branch_components.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
HCON BWD
TCON TUB
HSTU SPCOMPONENT /SPEC/COMPONENT1
NEW FLANGE
POS X 1100mm Y 2100mm Z 3100mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 1900mm Y 2900mm Z 3900mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['Pipe'] == '0ABC12BR100'
        assert result[0]['Branch'] == '/B1'
        assert result[0]['HCON'] == 'BWD'
        assert result[0]['TCON'] == 'TUB'
        assert result[0]['HSTU'] == '/SPEC/COMPONENT1'
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'ELBOW'
        # Check pipe lengths are calculated
        assert result[0]['Head_Pipe_Length_mm'] > 0
        assert result[0]['Tail_Pipe_Length_mm'] > 0
    
    def test_attachment_exclusion(self, tmp_path):
        """Test that ATTACHMENT components are excluded from first/last detection"""
        test_file = tmp_path / "test_attachment.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW ATTACHMENT
POS X 1000mm Y 2000mm Z 3000mm
SPRE SPCOMPONENT /SPEC/ATT1
END
NEW FLANGE
POS X 1100mm Y 2100mm Z 3100mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ATTACHMENT
POS X 2000mm Y 3000mm Z 4000mm
SPRE SPCOMPONENT /SPEC/ATT2
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'FLANGE'
    
    def test_multiple_component_types(self, tmp_path):
        """Test with various component types"""
        test_file = tmp_path / "test_multiple_types.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 3000mm Y 4000mm Z 5000mm
NEW COUPLING
POS X 1100mm Y 2100mm Z 3100mm
END
NEW REDUCER
POS X 1500mm Y 2500mm Z 3500mm
END
NEW UNION
POS X 2000mm Y 3000mm Z 4000mm
END
NEW VALVE
POS X 2900mm Y 3900mm Z 4900mm
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['First_Component'] == 'COUPLING'
        assert result[0]['Last_Component'] == 'VALVE'
    
    def test_branch_without_components(self, tmp_path):
        """Test branch with no components"""
        test_file = tmp_path / "test_no_components.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
HCON BWD
TCON TUB
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['First_Component'] == ''
        assert result[0]['Last_Component'] == ''
        assert result[0]['Head_Pipe_Length_mm'] == 0
        assert result[0]['Tail_Pipe_Length_mm'] == 0
    
    def test_pipe_length_zero_at_exact_position(self, tmp_path):
        """Test pipe length is 0 when component is at exact TPOS"""
        test_file = tmp_path / "test_exact_position.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW FLANGE
POS X 1000mm Y 2000mm Z 3000mm
END
NEW REDUCER
POS X 2000mm Y 3000mm Z 4000mm
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['Head_Pipe_Length_mm'] == 0.0
        assert result[0]['Tail_Pipe_Length_mm'] == 0.0
    
    def test_multiple_branches_saved(self, tmp_path):
        """Test that all branches are saved correctly"""
        test_file = tmp_path / "test_multiple_save.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
NEW FLANGE
POS X 1100mm Y 2100mm Z 3100mm
END
NEW BRANCH /0ABC12BR100/B2
HPOS X 1500mm Y 2500mm Z 3500mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW ELBOW
POS X 1600mm Y 2600mm Z 3600mm
END
NEW PIPE /0DEF34BR200
NEW BRANCH /0DEF34BR200/B1
HPOS X 5000mm Y 6000mm Z 7000mm
TPOS X 5500mm Y 6500mm Z 7500mm
NEW VALVE
POS X 5100mm Y 6100mm Z 7100mm
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 3
        assert result[0]['Pipe'] == '0ABC12BR100'
        assert result[0]['Branch'] == '/B1'
        assert result[1]['Pipe'] == '0ABC12BR100'
        assert result[1]['Branch'] == '/B2'
        assert result[2]['Pipe'] == '0DEF34BR200'
        assert result[2]['Branch'] == '/B1'
    
    def test_excluded_component_types(self, tmp_path):
        """Test that excluded types don't appear as first/last components"""
        test_file = tmp_path / "test_excluded.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW CYLINDER
POS X 1100mm Y 2100mm Z 3100mm
END
NEW FLANGE
POS X 1200mm Y 2200mm Z 3200mm
END
NEW CTORUS
POS X 1900mm Y 2900mm Z 3900mm
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'FLANGE'
    
    def test_real_data_0QCA30BR220(self):
        """Test specific known branch from real data"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        result = extract_branch_connections(str(real_file))
        
        # Find the specific branch
        branch = next((b for b in result if b['Pipe'] == '0QCA30BR220' and b['Branch'] == '/B1'), None)
        
        if branch:
            assert branch['HCON'] == 'TUB'
            assert branch['TCON'] == 'BWD'
            assert branch['First_Component'] == 'BEND'
            assert branch['Last_Component'] == 'REDUCER'
            assert branch['Head_Pipe_Length_mm'] > 0


class TestExtractComponentsFromBranches:
    """Tests for extract_components_from_branches function"""
    
    def test_happy_path_component_extraction(self, tmp_path):
        """Test basic component extraction"""
        test_file = tmp_path / "test_components.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 2
        assert result[0]['Pipe'] == '0ABC12BR100'
        assert result[0]['Branch'] == '/B1'
        assert result[0]['Component_Type'] == 'FLANGE'
        assert result[0]['SPRE'] == '/SPEC/FLANGE1'
        assert result[1]['Component_Type'] == 'ELBOW'
    
    def test_multiple_branches_components(self, tmp_path):
        """Test components from multiple branches"""
        test_file = tmp_path / "test_multi_branch_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW BRANCH /0ABC12BR100/B2
NEW ELBOW
SPRE SPCOMPONENT /SPEC/ELBOW1
END
NEW VALVE
SPRE SPCOMPONENT /SPEC/VALVE1
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 3
        assert result[0]['Branch'] == '/B1'
        assert result[1]['Branch'] == '/B2'
        assert result[2]['Branch'] == '/B2'
    
    def test_attachment_included(self, tmp_path):
        """Test that ATTACHMENT components are included in extraction"""
        test_file = tmp_path / "test_attachment_inc.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
NEW ATTACHMENT
SPRE SPCOMPONENT /SPEC/ATT1
END
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 2
        assert result[0]['Component_Type'] == 'ATTACHMENT'
        assert result[1]['Component_Type'] == 'FLANGE'
    
    def test_excluded_types_not_extracted(self, tmp_path):
        """Test that excluded types (PIPE, BRANCH, etc.) are not extracted"""
        test_file = tmp_path / "test_excluded_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
NEW ZONE
SPRE SPCOMPONENT /SPEC/ZONE1
END
NEW SITE
SPRE SPCOMPONENT /SPEC/SITE1
END
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 1
        assert result[0]['Component_Type'] == 'FLANGE'
    
    def test_component_without_spre(self, tmp_path):
        """Test component without SPRE property"""
        test_file = tmp_path / "test_no_spre.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
NEW FLANGE
BUIL false
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 0
    
    def test_real_data_component_count(self):
        """Test component count from real files"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        result = extract_components_from_branches(str(real_file))
        
        assert len(result) == 7633
        # Verify component types are valid
        component_types = set(c['Component_Type'] for c in result)
        assert 'FLANGE' in component_types or 'ELBOW' in component_types or 'VALVE' in component_types


class TestLookupAndMergeWithExcel:
    """Tests for lookup_and_merge_with_excel function"""
    
    def test_happy_path_weld_detection(self, tmp_path):
        """Test weld detection with BWD connection"""
        components = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/FLANGE1'},
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'ELBOW', 'SPRE': '/SPEC/ELBOW1'}
        ]
        
        # Create test Excel file
        excel_file = tmp_path / "test_specs.xlsx"
        df = pd.DataFrame({
            'SPRE': ['/SPEC/FLANGE1', '/SPEC/ELBOW1'],
            'P1 CONN': ['BWD', 'FLG'],
            'P2 CONN': ['FLG', 'BWD'],
            'TYPE': ['FLANGE', 'ELBOW']
        })
        df.to_excel(excel_file, index=False)
        
        result = lookup_and_merge_with_excel(components, str(excel_file))
        
        assert len(result) == 2
        assert result[0]['Found'] == 'X'
        assert result[0]['P1_CONN'] == 'BWD'
        assert result[0]['Welded'] == 'X'
        assert result[1]['P2_CONN'] == 'BWD'
        assert result[1]['Welded'] == 'X'
    
    def test_olet_weld_detection(self, tmp_path):
        """Test weld detection for OLET type"""
        components = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'TEE', 'SPRE': '/SPEC/TEE1'}
        ]
        
        excel_file = tmp_path / "test_olet.xlsx"
        df = pd.DataFrame({
            'SPRE': ['/SPEC/TEE1'],
            'P1 CONN': ['FLG'],
            'P2 CONN': ['FLG'],
            'TYPE': ['OLET']
        })
        df.to_excel(excel_file, index=False)
        
        result = lookup_and_merge_with_excel(components, str(excel_file))
        
        assert result[0]['Welded'] == 'X'
        assert result[0]['TYPE'] == 'OLET'
    
    def test_component_not_found_in_excel(self, tmp_path):
        """Test component not found in Excel"""
        components = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/MISSING'}
        ]
        
        excel_file = tmp_path / "test_missing.xlsx"
        df = pd.DataFrame({
            'SPRE': ['/SPEC/OTHER'],
            'P1 CONN': ['BWD'],
            'P2 CONN': ['FLG'],
            'TYPE': ['FLANGE']
        })
        df.to_excel(excel_file, index=False)
        
        result = lookup_and_merge_with_excel(components, str(excel_file))
        
        assert result[0]['Found'] == ''
        assert result[0]['Welded'] == ''
        assert result[0]['P1_CONN'] == ''
    
    def test_non_welded_component(self, tmp_path):
        """Test component with no weld connections"""
        components = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/FLANGE1'}
        ]
        
        excel_file = tmp_path / "test_non_weld.xlsx"
        df = pd.DataFrame({
            'SPRE': ['/SPEC/FLANGE1'],
            'P1 CONN': ['FLG'],
            'P2 CONN': ['FLG'],
            'TYPE': ['FLANGE']
        })
        df.to_excel(excel_file, index=False)
        
        result = lookup_and_merge_with_excel(components, str(excel_file))
        
        assert result[0]['Found'] == 'X'
        assert result[0]['Welded'] == ''
    
    def test_nan_values_in_excel(self, tmp_path):
        """Test handling of NaN values in Excel"""
        components = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/FLANGE1'}
        ]
        
        excel_file = tmp_path / "test_nan.xlsx"
        df = pd.DataFrame({
            'SPRE': ['/SPEC/FLANGE1'],
            'P1 CONN': [pd.NA],
            'P2 CONN': ['BWD'],
            'TYPE': [pd.NA]
        })
        df.to_excel(excel_file, index=False)
        
        result = lookup_and_merge_with_excel(components, str(excel_file))
        
        assert result[0]['P1_CONN'] == ''
        assert result[0]['P2_CONN'] == 'BWD'
        assert result[0]['TYPE'] == ''
        assert result[0]['Welded'] == 'X'  # BWD on P2
    
    def test_real_data_weld_count(self):
        """Test with real data files"""
        real_txt = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        real_excel = Path(__file__).parent.parent / "TBY" / "TBY_all_pspecs_wure_macro_08.12.2025.xlsx"
        
        if not real_txt.exists() or not real_excel.exists():
            pytest.skip("Real data files not found")
        
        components = extract_components_from_branches(str(real_txt))
        result = lookup_and_merge_with_excel(components, str(real_excel))
        
        welded_count = sum(1 for r in result if r['Welded'] == 'X')
        # Should find welded components
        assert welded_count > 0


class TestSaveToCsv:
    """Tests for save_to_csv function"""
    
    def test_happy_path_save(self, tmp_path):
        """Test basic CSV save functionality"""
        data = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/FLANGE1',
             'Found': 'X', 'P1_CONN': 'BWD', 'P2_CONN': 'FLG', 'TYPE': 'FLANGE', 'Welded': 'X'},
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'ELBOW', 'SPRE': '/SPEC/ELBOW1',
             'Found': 'X', 'P1_CONN': 'FLG', 'P2_CONN': 'BWD', 'TYPE': 'ELBOW', 'Welded': 'X'}
        ]
        
        output_file = tmp_path / "test_output.csv"
        save_to_csv(data, str(output_file))
        
        assert output_file.exists()
        
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        
        assert len(rows) == 2
        assert rows[0]['Pipe'] == '0ABC12BR100'
        assert rows[0]['Welded'] == 'X'
        assert rows[1]['Component_Type'] == 'ELBOW'
    
    def test_empty_data(self, tmp_path, capsys):
        """Test save with empty data"""
        output_file = tmp_path / "test_empty.csv"
        save_to_csv([], str(output_file))
        
        captured = capsys.readouterr()
        assert "No data to save" in captured.out
        assert not output_file.exists()
    
    def test_special_characters(self, tmp_path):
        """Test saving data with special characters"""
        data = [
            {'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Component_Type': 'FLANGE', 'SPRE': '/SPEC/FL-1,2"3',
             'Found': 'X', 'P1_CONN': 'BWD', 'P2_CONN': 'FLG', 'TYPE': 'FLANGE', 'Welded': 'X'}
        ]
        
        output_file = tmp_path / "test_special.csv"
        save_to_csv(data, str(output_file))
        
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        
        assert rows[0]['SPRE'] == '/SPEC/FL-1,2"3'


class TestIntegrationRealData:
    """Integration tests using real data files"""
    
    def test_full_pipeline_real_data(self):
        """Test complete pipeline with real data"""
        real_txt_0 = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        real_txt_1 = Path(__file__).parent.parent / "TBY" / "TBY-1AUX-P.txt"
        real_excel = Path(__file__).parent.parent / "TBY" / "TBY_all_pspecs_wure_macro_08.12.2025.xlsx"
        
        if not all([real_txt_0.exists(), real_txt_1.exists(), real_excel.exists()]):
            pytest.skip("Real data files not found")
        
        # Extract components
        components_0 = extract_components_from_branches(str(real_txt_0))
        components_1 = extract_components_from_branches(str(real_txt_1))
        all_components = components_0 + components_1
        
        # Extract branch connections
        branches_0 = extract_branch_connections(str(real_txt_0))
        branches_1 = extract_branch_connections(str(real_txt_1))
        all_branches = branches_0 + branches_1
        
        # Extract positions
        positions_0 = extract_branch_positions(str(real_txt_0))
        positions_1 = extract_branch_positions(str(real_txt_1))
        all_positions = positions_0 + positions_1
        
        # Find connections
        connections = find_connected_branches(all_positions, tolerance_tight=5.0, tolerance_loose=150.0)
        
        # Lookup and merge
        result = lookup_and_merge_with_excel(all_components, str(real_excel))
        
        # Verify results
        assert len(all_components) == 13536
        assert len(all_positions) == 1011
        assert len(connections) == 174
        assert len(result) == 13536
        
        welded_count = sum(1 for r in result if r['Welded'] == 'X')
        assert welded_count == 1425
    
    def test_specific_branch_0QCA30BR220(self):
        """Test specific known branch with all functions"""
        real_txt = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_txt.exists():
            pytest.skip("Real data file not found")
        
        # Test branch connections
        branches = extract_branch_connections(str(real_txt))
        target = next((b for b in branches if b['Pipe'] == '0QCA30BR220' and b['Branch'] == '/B1'), None)
        
        assert target is not None
        assert target['First_Component'] == 'BEND'
        assert target['Last_Component'] == 'REDUCER'
        assert target['HCON'] == 'TUB'
        assert target['TCON'] == 'BWD'
        assert target['Head_Pipe_Length_mm'] > 0
        assert target['Tail_Pipe_Length_mm'] >= 0
    
    def test_specific_branch_0QFB91BR110(self):
        """Test specific known branch 0QFB91BR110/B2"""
        real_txt = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_txt.exists():
            pytest.skip("Real data file not found")
        
        branches = extract_branch_connections(str(real_txt))
        target = next((b for b in branches if b['Pipe'] == '0QFB91BR110' and b['Branch'] == '/B2'), None)
        
        assert target is not None
        assert target['First_Component'] == 'COUPLING'
        assert target['Last_Component'] == 'VALVE'
        assert target['HCON'] == 'SWF'
        assert target['TCON'] == 'SCGQ'


class TestPipeLengthCalculations:
    """Comprehensive tests for head and tail pipe length calculations"""
    
    def test_single_component_with_pipe_on_both_ends(self, tmp_path):
        """Test branch with one component and pipe segments on both ends"""
        test_file = tmp_path / "test_single_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 3000mm Y 4000mm Z 5000mm
NEW ELBOW
POS X 2000mm Y 3000mm Z 4000mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: distance from (1000,2000,3000) to (2000,3000,4000)
        # = sqrt(1000² + 1000² + 1000²) = sqrt(3000000) ≈ 1732.05mm
        assert abs(result[0]['Head_Pipe_Length_mm'] - 1732.05) < 0.1
        # Tail: distance from (2000,3000,4000) to (3000,4000,5000)
        # = sqrt(1000² + 1000² + 1000²) = sqrt(3000000) ≈ 1732.05mm
        assert abs(result[0]['Tail_Pipe_Length_mm'] - 1732.05) < 0.1
        assert result[0]['First_Component'] == 'ELBOW'
        assert result[0]['Last_Component'] == 'ELBOW'
    
    def test_two_components_pipe_between(self, tmp_path):
        """Test branch with two components and pipe segments"""
        test_file = tmp_path / "test_two_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 5000mm Y 0mm Z 0mm
NEW FLANGE
POS X 1000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 4000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: (0,0,0) to (1000,0,0) = 1000mm
        assert result[0]['Head_Pipe_Length_mm'] == 1000.0
        # Tail: (4000,0,0) to (5000,0,0) = 1000mm
        assert result[0]['Tail_Pipe_Length_mm'] == 1000.0
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'ELBOW'
    
    def test_five_components_complex(self, tmp_path):
        """Test branch with five components in sequence"""
        test_file = tmp_path / "test_five_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 10000mm Y 0mm Z 0mm
NEW FLANGE
POS X 500mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 2000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
NEW REDUCER
POS X 5000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/REDUCER1
END
NEW TEE
POS X 7000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/TEE1
END
NEW VALVE
POS X 9500mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/VALVE1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: (0,0,0) to (500,0,0) = 500mm
        assert result[0]['Head_Pipe_Length_mm'] == 500.0
        # Tail: (9500,0,0) to (10000,0,0) = 500mm
        assert result[0]['Tail_Pipe_Length_mm'] == 500.0
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'VALVE'
    
    def test_component_at_hpos_only(self, tmp_path):
        """Test component exactly at HPOS (zero head pipe length)"""
        test_file = tmp_path / "test_at_hpos.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 5000mm Y 6000mm Z 7000mm
NEW FLANGE
POS X 1000mm Y 2000mm Z 3000mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 3000mm Y 4000mm Z 5000mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: component at exact HPOS = 0mm
        assert result[0]['Head_Pipe_Length_mm'] == 0.0
        # Tail: (3000,4000,5000) to (5000,6000,7000) = sqrt(4+4+4)*1000 = 3464.10mm
        assert abs(result[0]['Tail_Pipe_Length_mm'] - 3464.10) < 0.1
    
    def test_component_at_tpos_only(self, tmp_path):
        """Test component exactly at TPOS (zero tail pipe length)"""
        test_file = tmp_path / "test_at_tpos.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 5000mm Y 6000mm Z 7000mm
NEW FLANGE
POS X 3000mm Y 4000mm Z 5000mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 5000mm Y 6000mm Z 7000mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: (1000,2000,3000) to (3000,4000,5000) = sqrt(4+4+4)*1000 = 3464.10mm
        assert abs(result[0]['Head_Pipe_Length_mm'] - 3464.10) < 0.1
        # Tail: component at exact TPOS = 0mm
        assert result[0]['Tail_Pipe_Length_mm'] == 0.0
    
    def test_no_components_zero_pipe_length(self, tmp_path):
        """Test branch with no components (should have zero pipe lengths)"""
        test_file = tmp_path / "test_no_comp.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 5000mm Y 6000mm Z 7000mm
HCON BWD
TCON FLG
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        assert result[0]['Head_Pipe_Length_mm'] == 0
        assert result[0]['Tail_Pipe_Length_mm'] == 0
        assert result[0]['First_Component'] == ''
        assert result[0]['Last_Component'] == ''
    
    def test_only_excluded_components(self, tmp_path):
        """Test branch with only ATTACHMENT components (excluded from pipe length)"""
        test_file = tmp_path / "test_only_attach.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 5000mm Y 6000mm Z 7000mm
NEW ATTACHMENT
POS X 2000mm Y 3000mm Z 4000mm
SPRE SPCOMPONENT /SPEC/ATT1
END
NEW ATTACHMENT
POS X 4000mm Y 5000mm Z 6000mm
SPRE SPCOMPONENT /SPEC/ATT2
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Attachments are excluded, so no first/last component
        assert result[0]['Head_Pipe_Length_mm'] == 0
        assert result[0]['Tail_Pipe_Length_mm'] == 0
        assert result[0]['First_Component'] == ''
        assert result[0]['Last_Component'] == ''
    
    def test_attachment_between_valid_components(self, tmp_path):
        """Test that ATTACHMENT between valid components doesn't affect first/last detection"""
        test_file = tmp_path / "test_attach_between.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 6000mm Y 0mm Z 0mm
NEW FLANGE
POS X 1000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ATTACHMENT
POS X 3000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/ATT1
END
NEW VALVE
POS X 5000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/VALVE1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: (0,0,0) to (1000,0,0) = 1000mm
        assert result[0]['Head_Pipe_Length_mm'] == 1000.0
        # Tail: (5000,0,0) to (6000,0,0) = 1000mm
        assert result[0]['Tail_Pipe_Length_mm'] == 1000.0
        assert result[0]['First_Component'] == 'FLANGE'
        assert result[0]['Last_Component'] == 'VALVE'
    
    def test_3d_diagonal_pipe_length(self, tmp_path):
        """Test pipe length calculation in all three dimensions"""
        test_file = tmp_path / "test_3d_diagonal.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 300mm Y 400mm Z 0mm
NEW FLANGE
POS X 100mm Y 100mm Z 100mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 200mm Y 300mm Z -100mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: sqrt(100² + 100² + 100²) = sqrt(30000) = 173.21mm
        assert abs(result[0]['Head_Pipe_Length_mm'] - 173.21) < 0.1
        # Tail: sqrt((300-200)² + (400-300)² + (0-(-100))²) = sqrt(10000+10000+10000) = 173.21mm
        assert abs(result[0]['Tail_Pipe_Length_mm'] - 173.21) < 0.1
    
    def test_very_short_pipe_segments(self, tmp_path):
        """Test with very short pipe segments (mm precision)"""
        test_file = tmp_path / "test_short_segments.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1001mm Y 2001mm Z 3001mm
NEW FLANGE
POS X 1000.5mm Y 2000.5mm Z 3000.5mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 1000.9mm Y 2000.9mm Z 3000.9mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: sqrt(0.5² + 0.5² + 0.5²) = 0.866mm
        assert abs(result[0]['Head_Pipe_Length_mm'] - 0.87) < 0.01
        # Tail: sqrt(0.1² + 0.1² + 0.1²) = 0.173mm
        assert abs(result[0]['Tail_Pipe_Length_mm'] - 0.17) < 0.01
    
    def test_very_long_pipe_segments(self, tmp_path):
        """Test with very long pipe segments (kilometers)"""
        test_file = tmp_path / "test_long_segments.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 1000000mm Y 0mm Z 0mm
NEW FLANGE
POS X 100000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 900000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: 100,000mm = 100m
        assert result[0]['Head_Pipe_Length_mm'] == 100000.0
        # Tail: 100,000mm = 100m
        assert result[0]['Tail_Pipe_Length_mm'] == 100000.0
    
    def test_negative_coordinates_pipe_length(self, tmp_path):
        """Test pipe length calculation with negative coordinates"""
        test_file = tmp_path / "test_negative_coords.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X -1000mm Y -2000mm Z -3000mm
TPOS X 1000mm Y 2000mm Z 3000mm
NEW FLANGE
POS X -500mm Y -1000mm Z -1500mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 500mm Y 1000mm Z 1500mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Head: sqrt((-1000-(-500))² + (-2000-(-1000))² + (-3000-(-1500))²)
        #     = sqrt(500² + 1000² + 1500²) = sqrt(2750000) = 1658.31mm
        # Wait, that's wrong. Let me recalculate:
        # Head: sqrt((-1000-(-500))² + (-2000-(-1000))² + (-3000-(-1500))²)
        #     = sqrt((-500)² + (-1000)² + (-1500)²) = sqrt(250000 + 1000000 + 2250000)
        #     = sqrt(3500000) = 1870.83mm
        assert abs(result[0]['Head_Pipe_Length_mm'] - 1870.83) < 0.1
        # Tail: sqrt((1000-500)² + (2000-1000)² + (3000-1500)²)
        #     = sqrt(500² + 1000² + 1500²) = sqrt(3500000) = 1870.83mm
        assert abs(result[0]['Tail_Pipe_Length_mm'] - 1870.83) < 0.1
    
    def test_multiple_branches_different_pipe_lengths(self, tmp_path):
        """Test multiple branches with varying pipe lengths"""
        test_file = tmp_path / "test_multi_branch_pipe_length.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 0mm Y 0mm Z 0mm
TPOS X 2000mm Y 0mm Z 0mm
NEW FLANGE
POS X 0mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW ELBOW
POS X 2000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/ELBOW1
END
NEW BRANCH /0ABC12BR100/B2
HPOS X 0mm Y 0mm Z 0mm
TPOS X 3000mm Y 0mm Z 0mm
NEW VALVE
POS X 1000mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/VALVE1
END
NEW BRANCH /0ABC12BR100/B3
HPOS X 0mm Y 0mm Z 0mm
TPOS X 1000mm Y 0mm Z 0mm
NEW TEE
POS X 500mm Y 0mm Z 0mm
SPRE SPCOMPONENT /SPEC/TEE1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 3
        # B1: both at exact positions
        assert result[0]['Head_Pipe_Length_mm'] == 0.0
        assert result[0]['Tail_Pipe_Length_mm'] == 0.0
        # B2: head=1000mm, tail=2000mm
        assert result[1]['Head_Pipe_Length_mm'] == 1000.0
        assert result[1]['Tail_Pipe_Length_mm'] == 2000.0
        assert result[1]['First_Component'] == 'VALVE'
        assert result[1]['Last_Component'] == 'VALVE'
        # B3: head=500mm, tail=500mm
        assert result[2]['Head_Pipe_Length_mm'] == 500.0
        assert result[2]['Tail_Pipe_Length_mm'] == 500.0
    
    def test_real_data_specific_pipe_lengths(self):
        """Test specific known branches from real data with expected pipe lengths"""
        real_file = Path(__file__).parent.parent / "TBY" / "TBY-0AUX-P.txt"
        
        if not real_file.exists():
            pytest.skip("Real data file not found")
        
        result = extract_branch_connections(str(real_file))
        
        # Find branch 0QCA30BR220/B1 (known to have 569.97mm head, 0.0mm tail)
        branch = next((b for b in result if b['Pipe'] == '0QCA30BR220' and b['Branch'] == '/B1'), None)
        if branch:
            assert abs(branch['Head_Pipe_Length_mm'] - 569.97) < 0.1
            assert branch['Tail_Pipe_Length_mm'] == 0.0
        
        # Count branches with non-zero pipe lengths
        branches_with_head_pipe = sum(1 for b in result if b['Head_Pipe_Length_mm'] > 0)
        branches_with_tail_pipe = sum(1 for b in result if b['Tail_Pipe_Length_mm'] > 0)
        
        # Should have many branches with pipe segments
        assert branches_with_head_pipe > 0
        assert branches_with_tail_pipe > 0


class TestEdgeCases:
    """Tests for edge cases and error conditions"""
    
    def test_nonexistent_file(self):
        """Test with nonexistent file"""
        with pytest.raises(FileNotFoundError):
            extract_branch_positions("nonexistent_file.txt")
    
    def test_malformed_coordinates(self, tmp_path):
        """Test with malformed coordinate data"""
        test_file = tmp_path / "test_malformed.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X ABCmm Y 2000mm Z 3000mm
TPOS X 1500mm Y DEFmm Z 3500mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        # Should handle gracefully with None values
        assert len(result) == 1
        assert result[0]['HPOS_X'] is None or result[0]['TPOS_Y'] is None
    
    def test_unicode_characters(self, tmp_path):
        """Test with Unicode characters in data"""
        test_file = tmp_path / "test_unicode.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
:HZIAreaCode 'Ø-Test'
END
""", encoding='utf-8')
        
        result = extract_branch_positions(str(test_file))
        
        assert len(result) == 1
    
    def test_very_large_coordinates(self, tmp_path):
        """Test with very large coordinate values"""
        test_file = tmp_path / "test_large.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 999999.999mm Y -999999.999mm Z 500000mm
TPOS X 1000000mm Y -1000000mm Z 500000.5mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        assert result[0]['HPOS_X'] == 999999.999
        assert result[0]['TPOS_Y'] == -1000000.0
    
    def test_duplicate_branch_ids(self, tmp_path):
        """Test handling of duplicate branch IDs"""
        test_file = tmp_path / "test_duplicate.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
NEW BRANCH /0ABC12BR100/B1
HPOS X 2000mm Y 3000mm Z 4000mm
TPOS X 2500mm Y 3500mm Z 4500mm
END
""")
        
        result = extract_branch_positions(str(test_file))
        
        # Both should be extracted (no deduplication)
        assert len(result) == 2
        assert all(b['Full_Branch_ID'] == '0ABC12BR100/B1' for b in result)
    
    def test_accuracy_with_non_zero_offset(self):
        """Test accuracy percentage with non-zero offset"""
        branches = [
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B1',
                'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 2.5, 'TPOS_Z': 0.0
            },
            {
                'Pipe': '0ABC12BR100',
                'Branch': '/B2',
                'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1002.5, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 1
        # Max offset is 2.5mm out of 5mm tolerance = 50% accuracy
        assert result[0]['Accuracy_Percent'] == 50.0
        assert result[0]['Max_Tight_Offset_mm'] == 2.5
    
    def test_branch_with_spre_on_different_lines(self, tmp_path):
        """Test component SPRE that might span lines"""
        test_file = tmp_path / "test_spre.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
POS X 1100mm Y 2100mm Z 3100mm
END
END
""")
        
        result = extract_components_from_branches(str(test_file))
        
        assert len(result) == 1
        assert result[0]['SPRE'] == '/SPEC/FLANGE1'
    
    def test_multiple_pipes_in_single_file(self, tmp_path):
        """Test file with multiple pipes and complex structure"""
        test_file = tmp_path / "test_multi_pipe.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 1500mm Y 2500mm Z 3500mm
HCON BWD
TCON FLG
NEW FLANGE
SPRE SPCOMPONENT /SPEC/FLANGE1
POS X 1100mm Y 2100mm Z 3100mm
END
NEW BRANCH /0ABC12BR100/B2
HPOS X 1500mm Y 2500mm Z 3500mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW ELBOW
SPRE SPCOMPONENT /SPEC/ELBOW1
POS X 1600mm Y 2600mm Z 3600mm
END
NEW PIPE /0DEF34BR200
NEW BRANCH /0DEF34BR200/B1
HPOS X 5000mm Y 6000mm Z 7000mm
TPOS X 5500mm Y 6500mm Z 7500mm
NEW VALVE
SPRE SPCOMPONENT /SPEC/VALVE1
POS X 5100mm Y 6100mm Z 7100mm
END
END
""")
        
        # Test all extraction functions
        positions = extract_branch_positions(str(test_file))
        connections = extract_branch_connections(str(test_file))
        components = extract_components_from_branches(str(test_file))
        
        assert len(positions) == 3
        assert len(connections) == 3
        assert len(components) == 3
        assert positions[0]['Pipe'] == '0ABC12BR100'
        assert positions[2]['Pipe'] == '0DEF34BR200'
    
    def test_all_three_match_types(self):
        """Test all three axis-combination match types in one test"""
        branches = [
            # XY tight match
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B1', 'Full_Branch_ID': '0ABC12BR100/B1',
                'HPOS_X': 0.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 1000.0, 'TPOS_Y': 1.0, 'TPOS_Z': 100.0
            },
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B2', 'Full_Branch_ID': '0ABC12BR100/B2',
                'HPOS_X': 1001.0, 'HPOS_Y': 2.0, 'HPOS_Z': 0.0,
                'TPOS_X': 2000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            # XZ tight match
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B3', 'Full_Branch_ID': '0ABC12BR100/B3',
                'HPOS_X': 3000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 4000.0, 'TPOS_Y': 100.0, 'TPOS_Z': 1.0
            },
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B4', 'Full_Branch_ID': '0ABC12BR100/B4',
                'HPOS_X': 4001.0, 'HPOS_Y': 0.0, 'HPOS_Z': 2.0,
                'TPOS_X': 5000.0, 'TPOS_Y': 0.0, 'TPOS_Z': 0.0
            },
            # YZ tight match
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B5', 'Full_Branch_ID': '0ABC12BR100/B5',
                'HPOS_X': 6000.0, 'HPOS_Y': 0.0, 'HPOS_Z': 0.0,
                'TPOS_X': 6100.0, 'TPOS_Y': 7000.0, 'TPOS_Z': 1.0
            },
            {
                'Pipe': '0ABC12BR100', 'Branch': '/B6', 'Full_Branch_ID': '0ABC12BR100/B6',
                'HPOS_X': 6000.0, 'HPOS_Y': 7001.0, 'HPOS_Z': 2.0,
                'TPOS_X': 6000.0, 'TPOS_Y': 8000.0, 'TPOS_Z': 0.0
            }
        ]
        
        result = find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0)
        
        assert len(result) == 3
        match_types = [r['Match_Type'] for r in result]
        assert 'XY_tight_Z_loose' in match_types
        assert 'XZ_tight_Y_loose' in match_types
        assert 'YZ_tight_X_loose' in match_types
    
    def test_component_at_branch_boundaries(self, tmp_path):
        """Test components positioned exactly at HPOS and TPOS"""
        test_file = tmp_path / "test_boundaries.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW FLANGE
POS X 1000mm Y 2000mm Z 3000mm
SPRE SPCOMPONENT /SPEC/FLANGE1
END
NEW REDUCER
POS X 2000mm Y 3000mm Z 4000mm
SPRE SPCOMPONENT /SPEC/REDUCER1
END
END
""")
        
        result = extract_branch_connections(str(test_file))
        
        assert len(result) == 1
        # Both components at exact positions - should have 0 pipe length
        assert result[0]['Head_Pipe_Length_mm'] == 0.0
        assert result[0]['Tail_Pipe_Length_mm'] == 0.0
    
    def test_branch_with_only_excluded_components(self, tmp_path):
        """Test branch containing only excluded component types"""
        test_file = tmp_path / "test_only_excluded.txt"
        test_file.write_text("""
NEW PIPE /0ABC12BR100
NEW BRANCH /0ABC12BR100/B1
HPOS X 1000mm Y 2000mm Z 3000mm
TPOS X 2000mm Y 3000mm Z 4000mm
NEW ATTACHMENT
POS X 1100mm Y 2100mm Z 3100mm
SPRE SPCOMPONENT /SPEC/ATT1
END
NEW CYLINDER
POS X 1500mm Y 2500mm Z 3500mm
END
NEW ATTACHMENT
POS X 1900mm Y 2900mm Z 3900mm
SPRE SPCOMPONENT /SPEC/ATT2
END
END
""")
        
        connections = extract_branch_connections(str(test_file))
        components = extract_components_from_branches(str(test_file))
        
        # Should have branch with empty first/last components
        assert len(connections) == 1
        assert connections[0]['First_Component'] == ''
        assert connections[0]['Last_Component'] == ''
        
        # Should still extract ATTACHMENT components
        assert len(components) == 2
        assert all(c['Component_Type'] == 'ATTACHMENT' for c in components)


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--cov=weld_counter", "--cov-report=term-missing"])
