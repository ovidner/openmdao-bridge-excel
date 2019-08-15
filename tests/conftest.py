from hypothesis import settings


settings.register_profile("fast", max_examples=5, deadline=2000, derandomize=True)
settings.load_profile("fast")
