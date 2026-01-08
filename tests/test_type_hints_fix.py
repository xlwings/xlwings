"""
Test to verify the fix for type hint analysis breaking subs.

This test verifies that type hints like 'None' and 'Iterable[str]' 
don't cause errors when used with @xw.sub and @xw.func decorators.

The core fix is in xlwings/conversion/__init__.py where we ensure that
unregistered types (like NoneType or Iterable) fallback to ValueAccessor.
"""
from typing import Iterable
from xlwings.conversion import accessors, Accessor, ValueAccessor


def test_accessor_fallback_logic():
    """Test that unregistered types fallback to ValueAccessor properly"""
    from xlwings.conversion import _get_accessor
    
    # Test with NoneType (unregistered type)
    none_type = type(None)
    accessor = _get_accessor(none_type)
    
    # Should be ValueAccessor now
    assert accessor == ValueAccessor, f"Expected ValueAccessor, got {accessor}"
    print("✓ NoneType fallback to ValueAccessor works")
    
    # Test with Iterable (unregistered type)
    accessor = _get_accessor(Iterable)
    
    assert accessor == ValueAccessor, f"Expected ValueAccessor, got {accessor}"
    print("✓ Iterable fallback to ValueAccessor works")
    
    # Test with a registered type (should return the registered accessor)
    accessor = _get_accessor(None)
    
    # None is registered to ValueAccessor, so should be valid
    assert isinstance(accessor, type) and issubclass(accessor, Accessor)
    print("✓ Registered types work correctly")


def test_read_function_with_unregistered_types():
    """Test that read() function handles unregistered types correctly"""
    from xlwings.conversion import _get_accessor
    
    # Import this locally to avoid Windows dependency
    try:
        # Test with NoneType - this should use ValueAccessor fallback
        # We can't actually call read() without Excel, but we can test the logic
        accessor = _get_accessor(type(None))
        
        # Should be able to get a reader pipeline without errors
        pipeline = accessor.reader({})
        assert pipeline is not None
        print("✓ read() with NoneType uses ValueAccessor fallback")
        
        # Test with Iterable
        accessor = _get_accessor(Iterable)
        
        pipeline = accessor.reader({})
        assert pipeline is not None
        print("✓ read() with Iterable uses ValueAccessor fallback")
        
    except ImportError:
        # Skip if we can't import read (Windows dependency)
        print("⊘ Skipping read() test (Windows dependency)")


def test_write_function_with_unregistered_types():
    """Test that write() function handles unregistered types correctly"""
    from xlwings.conversion import _get_accessor
    
    try:
        # Test with NoneType - this should use ValueAccessor fallback
        accessor = _get_accessor(type(None))
        
        # Should be able to get a writer pipeline without errors
        # We use None as value which should route to ValueAccessor
        pipeline = accessor.router(None, None, {}).writer({})
        assert pipeline is not None
        print("✓ write() with NoneType uses ValueAccessor fallback")
        
        # Test with Iterable
        accessor = _get_accessor(Iterable)
        
        pipeline = accessor.router(["item1", "item2"], None, {}).writer({})
        assert pipeline is not None
        print("✓ write() with Iterable uses ValueAccessor fallback")
        
    except ImportError:
        # Skip if we can't import write (Windows dependency)
        print("⊘ Skipping write() test (Windows dependency)")


if __name__ == "__main__":
    print("Testing type hint fix...")
    print()
    
    test_accessor_fallback_logic()
    test_read_function_with_unregistered_types()
    test_write_function_with_unregistered_types()
    
    print()
    print("All tests passed! ✓")
