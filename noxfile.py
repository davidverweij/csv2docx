import nox

@nox.session(python=["3.8"])
def tests(session):
    """Run the test suite using poetry"""
    session.run("poetry", "install", external=True)
    session.run("pytest", "--cov")
