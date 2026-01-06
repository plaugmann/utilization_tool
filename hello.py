def greet(name=None):
    """Return a greeting. If `name` is provided, greet the name."""
    if name:
        return f"Hello, {name}!"
    return "Hello, World!"


def main():
    import sys
    name = sys.argv[1] if len(sys.argv) > 1 else None
    print(greet(name))


if __name__ == "__main__":
    main()
