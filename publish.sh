#!/bin/bash

# 检查是否提供了NPM令牌
if [ -z "$1" ]; then
  echo "请提供NPM令牌作为参数"
  echo "用法: ./publish.sh your-npm-token"
  exit 1
fi

# 设置NPM令牌
NPM_TOKEN=$1

# 创建临时.npmrc文件
echo "//registry.npmjs.org/:_authToken=$NPM_TOKEN" > .npmrc

# 发布包
npm publish --access public

# 删除临时.npmrc文件
rm .npmrc

echo "发布完成！"
