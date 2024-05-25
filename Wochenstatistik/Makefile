.PHONY: all build down clean fclean

DOCKER_IMAGE_NAME := wochenstatistik
DOCKER_BUILD_ARGS := --build-arg BUILD_CONFIGURATION=base

all: build

build:
	chmod +x initDocument.sh
	./initDocument.sh
	docker build $(DOCKER_BUILD_ARGS) -t $(DOCKER_IMAGE_NAME) .

run:
	docker run --env-file .env --rm -it $(DOCKER_IMAGE_NAME)

re: clean all

stop:
	docker stop

clean: stop
	docker image rm -f lef_image 2>/dev/null || true

fclean: clean
	docker volume rm -f db_volume >/dev/null 2>&1 || true
