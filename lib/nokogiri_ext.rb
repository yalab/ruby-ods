# -*- coding: utf-8 -*-
Nokogiri::XML::Element.module_eval do
  def add_element(name, attributes={})
    (prefix, name) = name.split(':') if name.include?(':')
    node = Nokogiri::XML::Node.new(name, self)
    attributes.each do |attr, val|
      node.set_attribute(attr, val)
    end
    ns = node.add_namespace_definition(prefix, Ods::NAMESPACES[prefix])
    node.namespace = ns
    self.add_child(node)
    node
  end

  def fetch(xpath)
    if node = self.xpath(xpath).first
      return node
    end

    return self.add_element(xpath) unless xpath.include?('/')

    xpath = xpath.split('/')
    last_path = xpath.pop
    fetch(xpath.join('/')).fetch(last_path)
  end
end
